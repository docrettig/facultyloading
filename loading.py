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

# Path(__file__) locates this Python file; with_name finds the bundled workbook beside it.
TEMPLATE = Path(__file__).with_name('template.xlsx')
OVERLOAD_TABLE = Path(__file__).with_name('overload_table.xlsx')
# A set holds the column names an HTML export must contain before we accept its table.
REQUIRED = {'Instructor Name','Instructor ID','Academic Period','Course ID','CRN','Title','Campus','Sub Academic Period','Enrollment','Credits','Contact Hours','Schedule','Status','Building','Section ID','Instructor Responsibility (%)'}

# HTMLParser calls the methods below as it encounters tags and text.
# This class turns HTML tables into nested lists: tables → rows → cell text.
class Tables(HTMLParser):
    # Initialize the parser. None means we are not currently collecting that kind of item.
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables=[]; self.table=None; self.row=None; self.cell=None
    # Opening table/tr/td tags begin a table, row, or cell. th is a header cell.
    def handle_starttag(self,tag,attrs):
        if tag=='table': self.table=[]
        elif tag=='tr' and self.table is not None:self.row=[]
        elif tag in ('th','td') and self.row is not None:self.cell=[]
        elif tag=='br' and self.cell is not None:self.cell.append(' ')
    # Text can arrive in pieces; append each piece to the current cell.
    def handle_data(self,data):
        if self.cell is not None:self.cell.append(data)
    # Closing tags finish cells, then rows, then tables. join/split cleans extra whitespace.
    def handle_endtag(self,tag):
        if tag in ('th','td') and self.cell is not None:
            self.row.append(' '.join(''.join(self.cell).split()));self.cell=None
        elif tag=='tr' and self.row is not None:
            if self.row:self.table.append(self.row)
            self.row=None
        elif tag=='table' and self.table is not None:
            self.tables.append(self.table);self.table=None

# Input: uploaded HTML bytes. Output: a list of dictionaries, one per schedule row.
def parse_export(blob):
    # decode turns bytes into text; feed makes the parser walk through that text.
    p=Tables();p.feed(blob.decode('utf-8-sig',errors='replace'))
    for table in p.tables:
        for i,head in enumerate(table):
            # Accept a table only when its header includes all the required schedule columns.
            if REQUIRED.issubset(set(head)):
                # zip pairs each column heading with its value; dict makes a named course record.
                # The list comprehension repeats this conversion for every row with the right cell count.
                result=[dict(zip(head,row)) for row in table[i+1:] if len(row)==len(head)]
                if not result: raise ValueError('The schedule table has no rows.')
                return result
    raise ValueError('No schedule table with the expected Ivy Tech columns was found.')

# Combine the two schedules, check their term, and group unique assignments by instructor.
def read_exports(first,second):
    a,b=parse_export(first),parse_export(second)
    # A set removes duplicate term codes. Both exports must refer to the same term.
    terms={r['Academic Period'] for r in a+b if r['Academic Period']}
    if len(terms)!=1:raise ValueError(f'Exports must contain one matching academic period; found: {", ".join(sorted(terms))}')
    records={}
    for r in a+b:
        if not r['Instructor Name'].strip() or r['Status'].strip().lower()=='cancelled':continue
        # This tuple identifies one instructor assignment to one CRN within one term.
        key=(r['Academic Period'],r['CRN'],r['Instructor ID'] or r['Instructor Name'].casefold())
        # Meeting rows can repeat an assignment. Keep one, but reject conflicting load data.
        if key in records:
            old=records[key]
            critical=('Course ID','Credits','Contact Hours','Enrollment','Campus','Academic Period','Instructor Responsibility (%)')
            if any(old[c]!=r[c] for c in critical):
                raise ValueError(f'Conflicting duplicate assignment for CRN {r["CRN"]}.')
            continue
        records[key]=r
    # defaultdict(list) automatically creates an empty course list for a new instructor.
    people=defaultdict(list)
    for row in records.values():
        key=(row['Instructor ID'] or row['Instructor Name'].casefold(),row['Instructor Name'])
        people[key].append(row)
    # Return two values: the one term code, and a dictionary sorted by instructor name.
    return terms.pop(),dict(sorted(people.items(),key=lambda kv:kv[0][1].casefold()))

# Convert numeric text from HTML into a number. Return None if conversion fails.
def number(value):
    try:return float(value.replace(',','')) if '.' in value else int(value.replace(',',''))
    except (ValueError,AttributeError):return None

# Excel percentages are fractions: the export’s 50 becomes 0.5, and 100 becomes 1.
def responsibility_fraction(row):
    percent=number(row.get('Instructor Responsibility (%)'))
    # Reject missing, infinite, or out-of-range percentages instead of guessing a load.
    if percent is None or not math.isfinite(percent) or not 0<=percent<=100:
        raise ValueError(f'Missing or invalid Instructor Responsibility (%) for CRN {row.get("CRN","?")}; expected 0 to 100.')
    return percent/100

# Infer the delivery mode from the schedule description and location fields.
def modality(row):
    text=(row['Long Title']+' '+row['Building']+' '+row['Section ID']).lower()
    if 'internet course' in text or 'online' in text:return 'Online'
    if 'telev.' in text or 'virtual' in text:return 'Virtual'
    if row['Days'] or row['Room']:return 'In person'
    return ''

# Return numeric limits in this order: maximum credits, minimum contacts, maximum contacts.
def load_limits(term, role):
    # [-2:] is a slice taking the final two characters: 10 Summer, 20 Fall, 30 Spring.
    suffix=str(term)[-2:]
    if suffix not in ('10','20','30') or not str(term).isdigit():
        raise ValueError(f'Unsupported academic period {term!r}; expected a code ending 10, 20, or 30.')
    if not role:raise ValueError('Select a role to calculate acceptable ranges.')
    # Suppress openpyxl warnings about unsupported Excel extensions while reading the template/table.
    with catch_warnings():
        simplefilter('ignore',UserWarning)
        table=load_workbook(OVERLOAD_TABLE,read_only=True,data_only=True).active
        # Find the overload-table row whose first column equals the selected roster role.
        matches=[r for r in table.iter_rows(min_row=2,values_only=True) if r[0]==role]
    if len(matches)!=1:raise ValueError(f'Role {role!r} is missing or duplicated in the overload table.')
    r=matches[0]
    # The conditional expression chooses Summer columns E–G or Fall/Spring columns B–D.
    max_credit,min_contact,max_contact=(r[4],r[5],r[6]) if suffix=='10' else (r[1],r[2],r[3])
    if any(v is None for v in (max_credit,min_contact,max_contact)):
        raise ValueError(f'The overload table has missing limits for {role}.')
    return max_credit,min_contact,max_contact

# Convert the numeric limits into the two labels displayed on the loading form.
def acceptable_ranges(term, role):
    max_credit,min_contact,max_contact=load_limits(term,role)
    return f'Up to {max_credit:g}',f'{min_contact:g}-{max_contact:g}'

class LoadOutOfTolerance(ValueError):
    """Valid load data for which no allowed overload selection meets both caps."""


def select_overloads(term, role, rows):
    """Choose the fewest non-P1 sections needed to meet both load caps.

    Among valid choices of the same size, preserve as much assigned credit and
    contact load as possible, preferring a remaining load within the contact
    range on ties. Raise an explicit error when two eligible courses cannot fit.
    """
    credit_cap,contact_min,contact_cap=load_limits(term,role)
    # Build a parallel list of (adjusted credits, adjusted contacts) for each course.
    amounts=[]
    for r in rows:
        credits,contact=number(r.get('Credits')),number(r.get('Contact Hours'))
        if credits is None or contact is None or credits<0 or contact<0:
            raise ValueError(f'Missing or invalid credits/contact hours for CRN {r.get("CRN","?")}.')
        # Multiply both loads by the instructor’s share before testing either overload limit.
        amounts.append((credits*responsibility_fraction(r),contact*responsibility_fraction(r)))
    # sum adds the loads; the underscore indicates a tuple value we do not need here.
    total_credit=sum(c for c,_ in amounts);total_contact=sum(h for _,h in amounts)
    # If neither maximum is exceeded, return an empty tuple: no courses are overloads.
    # 1e-9 is a tiny tolerance that avoids false failures from floating-point rounding.
    if total_credit<=credit_cap+1e-9 and total_contact<=contact_cap+1e-9:return ()
    # Store the positions of courses that may be overloads. P1 is excluded.
    eligible=[i for i,r in enumerate(rows) if r.get('Campus','').strip().upper()!='P1']
    # Try every one-course option first; only try two-course options if none of those works.
    for count in (1,2):
        solutions=[]
        # combinations generates each distinct selection of eligible course positions.
        for choice in combinations(eligible,count):
            # Calculate what the regular load would be after removing this proposed selection.
            remaining_credit=total_credit-sum(amounts[i][0] for i in choice)
            remaining_contact=total_contact-sum(amounts[i][1] for i in choice)
            if remaining_credit<=credit_cap+1e-9 and remaining_contact<=contact_cap+1e-9:
                # Prefer a load that remains within the table's minimum contact
                # range, then minimize unused capacity across both caps.
                penalty=max(0,contact_min-remaining_contact)
                unused=(credit_cap-remaining_credit)/max(credit_cap,1)+(contact_cap-remaining_contact)/max(contact_cap,1)
                # Use CRNs and course IDs as a stable final tie-breaker so reruns choose consistently.
                tie=tuple((rows[i].get('CRN',''),rows[i].get('Course ID','')) for i in choice)
                # The first tuple ranks the proposal; the second tuple holds its course positions.
                solutions.append(((penalty>1e-9,round(penalty,9),round(unused,9),tie),choice))
        # min compares ranking tuples from left to right. [1] retrieves the winning positions.
        if solutions:return min(solutions)[1]
    raise LoadOutOfTolerance(f'Cannot bring {total_credit:g} credit hours and {total_contact:g} weekly contact hours within the {role} limits ({credit_cap:g} credit, {contact_cap:g} contact) by marking at most two non-P1 courses as overloads. Review assignments manually.')

def expand_course_rows(sheet, count):
    """Insert additional course rows and move the footer with its formulas."""
    # Ten course slots already exist; extra is how many additional rows we need.
    extra=max(0,count-10)
    if not extra:return 0
    # Temporarily remove merged cells, then recreate them at their shifted positions below.
    merges=list(sheet.merged_cells.ranges)
    # Move footer merges by the same extra-row count; header merges keep their positions.
    for merged in merges:sheet.unmerge_cells(str(merged))
    # Save row heights and other settings because insertion does not update all metadata.
    dimensions={r:copy(d) for r,d in sheet.row_dimensions.items()}
    # Insert above the old totals row (19), pushing totals and signature fields downward.
    sheet.insert_rows(19,extra)
    for r in list(sheet.row_dimensions):
        if r>=19:del sheet.row_dimensions[r]
    for r,d in dimensions.items():
        target=r+extra if r>=19 else r
        d.index=target;sheet.row_dimensions[target]=d
    # Excel-style insertion moves footer references and extends course ranges.
    # This regular expression identifies cell references such as U19 and $P$19.
    ref=re.compile(r'(\$?[A-Z]{1,3})(\$?)([0-9]+)')
    # The regex calls this helper for each reference; only footer rows move downward.
    def shift(match):
        n=int(match.group(3))
        return match.group(1)+match.group(2)+str(n+extra if n>=19 else n)
    for row in sheet:
        for cell in row:
            if cell.data_type!='f':continue
            # Split the Excel formula into tokens so we change references without changing quoted text.
            tokens=Tokenizer(cell.value)
            for token in tokens.items:
                if token.type=='OPERAND' and token.subtype=='RANGE':
                    value=ref.sub(shift,token.value)
                    # Extend ranges such as I9:I18 through the new final course row.
                    value=re.sub(r'(\$?[A-Z]{1,3}\$?9:)(\$?[A-Z]{1,3})(\$?)18\b',lambda m:m.group(1)+m.group(2)+m.group(3)+str(18+extra),value)
                    token.value=value
            cell.value='='+''.join(t.value for t in tokens.items)
    # Give each inserted row the styles and formulas of the original last course row.
    for r in range(19,19+extra):
        sheet.row_dimensions[r]=copy(sheet.row_dimensions[18])
        sheet.row_dimensions[r].index=r
        for c in range(1,sheet.max_column+1):
            original=sheet.cell(18,c);target=sheet.cell(r,c)
            target._style=copy(original._style)
            # Translator adjusts relative formula references as Excel does when copying a formula.
            target.value=Translator(original.value,origin=original.coordinate).translate_formula(target.coordinate) if original.data_type=='f' else original.value
    # Move footer merges by the same extra-row count; header merges keep their positions.
    for merged in merges:
        a,b,c,d=merged.bounds
        sheet.merge_cells(start_row=b+extra if b>=19 else b,start_column=a,end_row=d+extra if d>=19 else d,end_column=c)
    return extra

# Create one faculty workbook and return its bytes for download or an email attachment.
# Arguments: term code, instructor name, course records, optional label, roster role, load fraction.
def generate(term,person,rows,semester_label='',role='',load_request=1,alerts=None):
    # Suppress openpyxl warnings about unsupported Excel extensions while reading the template/table.
    with catch_warnings():
        simplefilter('ignore',UserWarning)
        book=load_workbook(TEMPLATE)
    sheet=book['Loading Template']
    extra=expand_course_rows(sheet,len(rows))
    # end is the final course row. All totals below the course area also move by extra.
    end=18+extra
    sheet['K1']=semester_label.strip() or term
    sheet['K2']=person
    credit_range,contact_range=acceptable_ranges(term,role)
    # Sort by course ID and CRN so the output order is predictable.
    ordered=sorted(rows,key=lambda x:(x['Course ID'],x['CRN']))
    # A set makes testing whether a course position is an overload straightforward.
    # A tolerance problem is a warning, not a reason to suppress the document.
    # Invalid/missing source data still raises its normal validation error.
    tolerance_warning=''
    try:
        overloads=set(select_overloads(term,role,ordered))
    except LoadOutOfTolerance as error:
        overloads=set()
        tolerance_warning=f'OUT OF TOLERANCE — {person}: {error} No overloads selected.'
        if alerts is not None:alerts.append(tolerance_warning)
    # Remove conditional formatting from every sheet, as requested.
    for tab in book:
        tab.conditional_formatting._cf_rules.clear()
    if not 0<=load_request<=1:raise ValueError('Load request must be between 0 and 100%.')
    sheet['K4']=role
    # An f-string builds the cell address. U23 becomes U35 when 12 rows are inserted.
    sheet[f'U{23+extra}']=load_request
    # An f-string builds the cell address. U23 becomes U35 when 12 rows are inserted.
    sheet[f'U{23+extra}'].number_format='0%' if float(load_request*100).is_integer() else '0.00%'
    # Academic Success allowance: Dean/Adjunct = 0; other roles = 6.25 Summer or 8 Fall/Spring.
    sheet[f'O{19+extra}']=0 if role in ('Dean','Adjunct') else (6.25 if str(term).endswith('10') else 8)
    sheet[f'U{21+extra}']=credit_range
    sheet[f'V{22+extra}']=contact_range
    sheet[f'A{27+extra}']='COMMENTS: Review suggested overloads, clinical, externship, simulation and academic success hours, contract percentage, and role before approval.'
    # SUMIF subtracts marked overload rows from the regular weekly contact total.
    if tolerance_warning:
        sheet[f'A{27+extra}']=tolerance_warning
    sheet[f'U{20+extra}']=f'=SUM(U9:U{end})-SUMIF(P9:P{end},"Yes",U9:U{end})'
    sheet[f'V{21+extra}']=f'=SUM(V9:V{end})-SUMIF(P9:P{end},"Yes",V9:V{end})'
    # enumerate returns both a zero-based position and the course record at that position.
    for index,r in enumerate(ordered):
        # Course records start in Excel row 9, so add 9 to the Python position.
        n=index+9
        # Map export fields to Excel columns, then write each value into this course row.
        for col,val in {'A':r['Course ID'],'C':r['CRN'],'D':r['Title'],'E':r['Campus'], 'F':r['Sub Academic Period'] or r['Part of Term Code'], 'G':modality(r),'H':number(r['Enrollment'])}.items():
            sheet[f'{col}{n}']=val
        # Store a numeric fraction in Q; the number format displays it as a percentage.
        sheet[f'Q{n}']=responsibility_fraction(r)
        sheet[f'Q{n}'].number_format='0%' if float(number(r['Instructor Responsibility (%)'])).is_integer() else '0.00%'
        # Write Excel formulas rather than fixed results, so editing Q recalculates credits and contacts.
        sheet[f'I{n}']=f'={number(r["Credits"])}*Q{n}'
        sheet[f'U{n}']=f'={number(r["Contact Hours"])}*Q{n}'
        sheet[f'P{n}']='Yes' if index in overloads else 'No'
        # The export has total contact hours only; it does not establish the breakdown
        # or contract assignment. Those cells remain editable for faculty review.
    # Make the loading form the selected tab when the instructor opens the workbook.
    book.active=book.sheetnames.index('Loading Template')
    # BytesIO acts like a file in memory. Save there and return the complete .xlsx contents.
    out=BytesIO();book.save(out);return out.getvalue()
