import re
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import streamlit as st
from loading import read_exports, generate
from faculty import read_roster, match_roster, email_draft

st.set_page_config(page_title='Faculty loading documents',layout='centered')
st.title('Faculty loading documents')
st.write('Upload both schedule exports and the current faculty roster. Only faculty listed in the roster are included.')
a=st.file_uploader('Schedule export 1 (.html)',type=['html','htm'],key='first')
b=st.file_uploader('Schedule export 2 (.html)',type=['html','htm'],key='second')
roster=st.file_uploader('Current faculty roster (.xlsx)',type=['xlsx'],key='roster')
if a and b and roster:
    try:
        term,people=read_exports(a.getvalue(),b.getvalue())
        entries,errors=read_roster(roster.getvalue(),term)
        matches,matching_errors=match_roster(entries,people)
        for error in errors+matching_errors:st.error(error)
        semester=st.text_input('Semester label',value=term)
        products=[]
        for entry,person,rows in matches:
            try:
                workbook=generate(term,person[1],rows,semester,entry['role'],entry['loading'])
                safe=re.sub(r'[^A-Za-z0-9_-]+','_',entry['name']).strip('_')
                # Instructor ID makes filenames distinct even when names sanitize alike.
                filename=f'{term}_{safe}_{person[0]}_Loading.xlsx'
                draft=email_draft(entry,semester or term,filename,workbook)
                products.append((entry,filename,workbook,draft))
            except ValueError as error:st.error(f'{entry["name"]}: {error}')
        st.caption(f'{len(products)} faculty documents generated for academic period {term}.')
        st.info('Email drafts are downloadable .eml files with the loading workbook attached. Nothing is sent automatically. Open drafts in a compatible email client, review, and send. Some clients import .eml as messages rather than editable drafts.')
        if products:
            selected=st.selectbox('Faculty document',range(len(products)),format_func=lambda i:products[i][0]['name'])
            entry,filename,workbook,draft=products[selected]
            st.write(f'Role: {entry["role"]} · Load request: {entry["loading"]:.0%}')
            st.write('To: '+', '.join(entry['to']))
            st.write('CC: '+', '.join(dict.fromkeys(entry['cc'])))
            st.download_button('Download loading document',data=workbook,file_name=filename,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.download_button('Download email draft with attachment',data=draft,file_name=filename[:-5]+'.eml',mime='message/rfc822')
            archive=BytesIO()
            with ZipFile(archive,'w',ZIP_DEFLATED) as z:
                for entry,filename,workbook,draft in products:
                    z.writestr(filename,workbook)
                    z.writestr(filename[:-5]+'.eml',draft)
            st.download_button('Download all documents and email drafts',data=archive.getvalue(),file_name=f'{term}_Faculty_Loading_Documents.zip',mime='application/zip')
    except (ValueError,KeyError) as error:st.error(str(error))
    except Exception:
        st.error('Unable to read the uploads. Check that the roster is a valid .xlsx workbook and the schedule exports are valid HTML files.')
