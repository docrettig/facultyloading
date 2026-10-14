"""Parse Ivy Tech schedule HTML and fill the supplied loading template."""
from __future__ import annotations
from collections import defaultdict
from html.parser import HTMLParser
from io import BytesIO
from itertools import combinations
from copy import copy
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from pathlib import Path
import re
import math
from warnings import catch_warnings, simplefilter
from openpyxl import load_workbook

TEMPLATE = Path(__file__).with_name('template.xlsx')
OVERLOAD_TABLE = Path(__file__).with_name('overload_table.xlsx')
REQUIRED = {'Instructor Name','Instructor ID','Academic Period','Course ID','CRN','Title','Campus','Sub Academic Period','Enrollment','Credits','Contact Hours','Schedule','Status','Building','Section ID','Instructor Responsibility (%)'}

class Tables(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables=[]; self.table=None; self.row=None; self.cell=None
    def handle_starttag(self,tag,attrs):
        if tag=='table': self.table=[]
        elif tag=='tr' and self.table is not None:self.row=[]
        elif tag in ('th','td') and self.row is not None:self.cell=[]
        elif tag=='br' and self.cell is not None:self.cell.append(' ')
    def handle_data(self,data):
        if self.cell is not None:self.cell.append(data)
    def handle_endtag(self,tag):
        if tag in ('th','td') and self.cell is not None:
            self.row.append(' '.join(''.join(self.cell).split()));self.cell=None
        elif tag=='tr' and self.row is not None:
            if self.row:self.table.append(self.row)
            self.row=None
        elif tag=='table' and self.table is not None:
            self.tables.append(self.table);self.table=None

def parse_export(blob):
    p=Tables();p.feed(blob.decode('utf-8-sig',errors='replace'))
    for table in p.tables:
        for i,head in enumerate(table):
            if REQUIRED.issubset(set(head)):
                result=[dict(zip(head,row)) for row in table[i+1:] if len(row)==len(head)]
                if not result: raise ValueError('The schedule table has no rows.')
                return result
    raise ValueError('No schedule table with the expected Ivy Tech columns was found.')

def read_exports(first,second):
    a,b=parse_export(first),parse_export(second)
    terms={r['Academic Period'] for r in a+b if r['Academic Period']}
    if len(terms)!=1:raise ValueError(f'Exports must contain one matching academic period; found: {", ".join(sorted(terms))}')
    records={}
    for r in a+b:
        if not r['Instructor Name'].strip() or r['Status'].strip().lower()=='cancelled':continue
        key=(r['Academic Period'],r['CRN'],r['Instructor ID'] or r['Instructor Name'].casefold())
        if key in records:
            old=records[key]
            critical=('Course ID','Credits','Contact Hours','Enrollment','Campus','Academic Period','Instructor Responsibility (%)')
            if any(old[c]!=r[c] for c in critical):
                raise ValueError(f'Conflicting duplicate assignment for CRN {r["CRN"]}.')
            continue
        records[key]=r
    people=defaultdict(list)
    for row in records.values():
        key=(row['Instructor ID'] or row['Instructor Name'].casefold(),row['Instructor Name'])
        people[key].append(row)
    return terms.pop(),dict(sorted(people.items(),key=lambda kv:kv[0][1].casefold()))

def number(value):
    try:return float(value.replace(',','')) if '.' in value else int(value.replace(',',''))
    except (ValueError,AttributeError):return None

def responsibility_fraction(row):
    percent=number(row.get('Instructor Responsibility (%)'))
    if percent is None or not math.isfinite(percent) or not 0<=percent<=100:
        raise ValueError(f'Missing or invalid Instructor Responsibility (%) for CRN {row.get("CRN","?")}; expected 0 to 100.')
    return percent/100

def modality(row):
    text=(row['Long Title']+' '+row['Building']+' '+row['Section ID']).lower()
    if 'internet course' in text or 'online' in text:return 'Online'
    if 'telev.' in text or 'virtual' in text:return 'Virtual'
    if row['Days'] or row['Room']:return 'In person'
    return ''

def load_limits(term, role):
    suffix=str(term)[-2:]
    if suffix not in ('10','20','30') or not str(term).isdigit():
        raise ValueError(f'Unsupported academic period {term!r}; expected a code ending 10, 20, or 30.')
    if not role:raise ValueError('Select a role to calculate acceptable ranges.')
    with catch_warnings():
        simplefilter('ignore',UserWarning)
        table=load_workbook(OVERLOAD_TABLE,read_only=True,data_only=True).active
        matches=[r for r in table.iter_rows(min_row=2,values_only=True) if r[0]==role]
    if len(matches)!=1:raise ValueError(f'Role {role!r} is missing or duplicated in the overload table.')
    r=matches[0]
    max_credit,min_contact,max_contact=(r[4],r[5],r[6]) if suffix=='10' else (r[1],r[2],r[3])
    if any(v is None for v in (max_credit,min_contact,max_contact)):
        raise ValueError(f'The overload table has missing limits for {role}.')
    return max_credit,min_contact,max_contact

def acceptable_ranges(term, role):
    max_credit,min_contact,max_contact=load_limits(term,role)
    return f'Up to {max_credit:g}',f'{min_contact:g}-{max_contact:g}'

def select_overloads(term, role, rows):
    """Choose the fewest non-P1 sections needed to meet both load caps.

    Among valid choices of the same size, preserve as much assigned credit and
    contact load as possible, preferring a remaining load within the contact
    range on ties. Raise an explicit error when two eligible courses cannot fit.
    """
    credit_cap,contact_min,contact_cap=load_limits(term,role)
    amounts=[]
    for r in rows:
        credits,contact=number(r.get('Credits')),number(r.get('Contact Hours'))
        if credits is None or contact is None or credits<0 or contact<0:
            raise ValueError(f'Missing or invalid credits/contact hours for CRN {r.get("CRN","?")}.')
        amounts.append((credits*responsibility_fraction(r),contact*responsibility_fraction(r)))
    total_credit=sum(c for c,_ in amounts);total_contact=sum(h for _,h in amounts)
    if total_credit<=credit_cap+1e-9 and total_contact<=contact_cap+1e-9:return ()
    eligible=[i for i,r in enumerate(rows) if r.get('Campus','').strip().upper()!='P1']
    for count in (1,2):
        solutions=[]
        for choice in combinations(eligible,count):
            remaining_credit=total_credit-sum(amounts[i][0] for i in choice)
            remaining_contact=total_contact-sum(amounts[i][1] for i in choice)
            if remaining_credit<=credit_cap+1e-9 and remaining_contact<=contact_cap+1e-9:
                # Prefer a load that remains within the table's minimum contact
                # range, then minimize unused capacity across both caps.
                penalty=max(0,contact_min-remaining_contact)
                unused=(credit_cap-remaining_credit)/max(credit_cap,1)+(contact_cap-remaining_contact)/max(contact_cap,1)
                tie=tuple((rows[i].get('CRN',''),rows[i].get('Course ID','')) for i in choice)
                solutions.append(((penalty>1e-9,round(penalty,9),round(unused,9),tie),choice))
        if solutions:return min(solutions)[1]
    raise ValueError(f'Cannot bring {total_credit:g} credit hours and {total_contact:g} weekly contact hours within the {role} limits ({credit_cap:g} credit, {contact_cap:g} contact) by marking at most two non-P1 courses as overloads. Review assignments manually.')

def expand_course_rows(sheet, count):
    """Insert additional course rows and move the footer with its formulas."""
    extra=max(0,count-10)
    if not extra:return 0
    merges=list(sheet.merged_cells.ranges)
    for merged in merges:sheet.unmerge_cells(str(merged))
    dimensions={r:copy(d) for r,d in sheet.row_dimensions.items()}
    sheet.insert_rows(19,extra)
    for r in list(sheet.row_dimensions):
        if r>=19:del sheet.row_dimensions[r]
    for r,d in dimensions.items():
        target=r+extra if r>=19 else r
        d.index=target;sheet.row_dimensions[target]=d
    # Excel-style insertion moves footer references and extends course ranges.
    ref=re.compile(r'(\$?[A-Z]{1,3})(\$?)([0-9]+)')
    def shift(match):
        n=int(match.group(3))
        return match.group(1)+match.group(2)+str(n+extra if n>=19 else n)
    for row in sheet:
        for cell in row:
            if cell.data_type!='f':continue
            tokens=Tokenizer(cell.value)
            for token in tokens.items:
                if token.type=='OPERAND' and token.subtype=='RANGE':
                    value=ref.sub(shift,token.value)
                    value=re.sub(r'(\$?[A-Z]{1,3}\$?9:)(\$?[A-Z]{1,3})(\$?)18\b',lambda m:m.group(1)+m.group(2)+m.group(3)+str(18+extra),value)
                    token.value=value
            cell.value='='+''.join(t.value for t in tokens.items)
    for r in range(19,19+extra):
        sheet.row_dimensions[r]=copy(sheet.row_dimensions[18])
        sheet.row_dimensions[r].index=r
        for c in range(1,sheet.max_column+1):
            original=sheet.cell(18,c);target=sheet.cell(r,c)
            target._style=copy(original._style)
            target.value=Translator(original.value,origin=original.coordinate).translate_formula(target.coordinate) if original.data_type=='f' else original.value
    for merged in merges:
        a,b,c,d=merged.bounds
        sheet.merge_cells(start_row=b+extra if b>=19 else b,start_column=a,end_row=d+extra if d>=19 else d,end_column=c)
    return extra

def generate(term,person,rows,semester_label='',role='',load_request=1):
    with catch_warnings():
        simplefilter('ignore',UserWarning)
        book=load_workbook(TEMPLATE)
    sheet=book['Loading Template']
    extra=expand_course_rows(sheet,len(rows))
    end=18+extra
    sheet['K1']=semester_label.strip() or term
    sheet['K2']=person
    credit_range,contact_range=acceptable_ranges(term,role)
    ordered=sorted(rows,key=lambda x:(x['Course ID'],x['CRN']))
    overloads=set(select_overloads(term,role,ordered))
    for tab in book:
        tab.conditional_formatting._cf_rules.clear()
    if not 0<=load_request<=1:raise ValueError('Load request must be between 0 and 100%.')
    sheet['K4']=role
    sheet[f'U{23+extra}']=load_request
    sheet[f'U{23+extra}'].number_format='0%' if float(load_request*100).is_integer() else '0.00%'
    sheet[f'O{19+extra}']=0 if role in ('Dean','Adjunct') else (6.25 if str(term).endswith('10') else 8)
    sheet[f'U{21+extra}']=credit_range
    sheet[f'V{22+extra}']=contact_range
    sheet[f'A{27+extra}']='COMMENTS: Review suggested overloads, clinical, externship, simulation and academic success hours, contract percentage, and role before approval.'
    sheet[f'U{20+extra}']=f'=SUM(U9:U{end})-SUMIF(P9:P{end},"Yes",U9:U{end})'
    sheet[f'V{21+extra}']=f'=SUM(V9:V{end})-SUMIF(P9:P{end},"Yes",V9:V{end})'
    for index,r in enumerate(ordered):
        n=index+9
        for col,val in {'A':r['Course ID'],'C':r['CRN'],'D':r['Title'],'E':r['Campus'], 'F':r['Sub Academic Period'] or r['Part of Term Code'], 'G':modality(r),'H':number(r['Enrollment'])}.items():
            sheet[f'{col}{n}']=val
        sheet[f'Q{n}']=responsibility_fraction(r)
        sheet[f'Q{n}'].number_format='0%' if float(number(r['Instructor Responsibility (%)'])).is_integer() else '0.00%'
        sheet[f'I{n}']=f'={number(r["Credits"])}*Q{n}'
        sheet[f'U{n}']=f'={number(r["Contact Hours"])}*Q{n}'
        sheet[f'P{n}']='Yes' if index in overloads else 'No'
        # The export has total contact hours only; it does not establish the breakdown
        # or contract assignment. Those cells remain editable for faculty review.
    book.active=book.sheetnames.index('Loading Template')
    out=BytesIO();book.save(out);return out.getvalue()
