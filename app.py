# Start with this file to understand the workflow. Streamlit builds the page as this script runs.
# Streamlit reruns the script when a user changes an input; uploads live in the browser session.
import re
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import streamlit as st
from loading import read_exports, generate
from faculty import read_roster, match_roster, email_draft

st.set_page_config(page_title='Faculty loading documents',layout='centered')
st.title('Faculty loading documents')
st.write('Upload both schedule exports and the current faculty roster. Only faculty listed in the roster are included.')
# These widgets return uploaded-file objects, or None until the assistant supplies a file.
a=st.file_uploader('Schedule export 1 (.html)',type=['html','htm'],key='first')
b=st.file_uploader('Schedule export 2 (.html)',type=['html','htm'],key='second')
roster=st.file_uploader('Current faculty roster (.xlsx)',type=['xlsx'],key='roster')
# Begin only after all three uploads are present.
if a and b and roster:
    try:
        # getvalue reads the uploaded bytes; unpack the parser’s two returned values.
        term,people=read_exports(a.getvalue(),b.getvalue())
        # Read only the faculty listed in the roster, then match them against both schedules.
        entries,errors=read_roster(roster.getvalue(),term)
        matches,matching_errors=match_roster(entries,people)
        for error in errors+matching_errors:st.error(error)
        semester=st.text_input('Semester label',value=term)
        # Collect successful outputs as (roster entry, filename, Excel bytes, email bytes) tuples.
        products=[]
        # Each match supplies the roster record, instructor identity, and their course list.
        for entry,person,rows in matches:
            try:
                # The role and Load Request come from the roster rather than manual webpage selections.
                # Collect load warnings while still creating the workbook and draft.
                alerts=[]
                workbook=generate(term,person[1],rows,semester,entry['role'],entry['loading'],alerts=alerts)
                for alert in alerts:st.warning(alert)
                # Replace filename punctuation/spaces with underscores using a regular expression.
                safe=re.sub(r'[^A-Za-z0-9_-]+','_',entry['name']).strip('_')
                # Instructor ID makes filenames distinct even when names sanitize alike.
                filename=f'{term}_{safe}_{person[0]}_Loading.xlsx'
                # semester or term uses the entered label when present, otherwise the exported term code.
                draft=email_draft(entry,semester or term,filename,workbook)
                products.append((entry,filename,workbook,draft))
            # Report errors per instructor while allowing other documents to finish.
            except ValueError as error:st.error(f'{entry["name"]}: {error}')
        st.caption(f'{len(products)} faculty documents generated for academic period {term}.')
        st.info('Email drafts are downloadable .eml files with the loading workbook attached. Nothing is sent automatically. Open drafts in a compatible email client, review, and send. Some clients import .eml as messages rather than editable drafts.')
        if products:
            # The widget stores an index; the lambda displays the faculty name for that index.
            selected=st.selectbox('Faculty document',range(len(products)),format_func=lambda i:products[i][0]['name'])
            # Unpack the selected result into its four named pieces.
            entry,filename,workbook,draft=products[selected]
            st.write(f'Role: {entry["role"]} · Load request: {entry["loading"]:.0%}')
            st.write('To: '+', '.join(entry['to']))
            st.write('CC: '+', '.join(dict.fromkeys(entry['cc'])))
            # Download buttons send existing bytes to the browser; they do not send email.
            st.download_button('Download loading document',data=workbook,file_name=filename,mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            # Download buttons send existing bytes to the browser; they do not send email.
            st.download_button('Download email draft with attachment',data=draft,file_name=filename[:-5]+'.eml',mime='message/rfc822')
            # Build a ZIP in memory containing every successful workbook and its email draft.
            archive=BytesIO()
            # with closes the ZIP correctly, finishing its contents before download.
            with ZipFile(archive,'w',ZIP_DEFLATED) as z:
                for entry,filename,workbook,draft in products:
                    z.writestr(filename,workbook)
                    # [:-5] removes the .xlsx extension so we can append .eml.
                    z.writestr(filename[:-5]+'.eml',draft)
            # Download buttons send existing bytes to the browser; they do not send email.
            st.download_button('Download all documents and email drafts',data=archive.getvalue(),file_name=f'{term}_Faculty_Loading_Documents.zip',mime='application/zip')
    # Handle validation problems that affect the entire upload set.
    except (ValueError,KeyError) as error:st.error(str(error))
    # Show a friendly message for unreadable/corrupt files rather than a raw Python traceback.
    except Exception:
        st.error('Unable to read the uploads. Check that the roster is a valid .xlsx workbook and the schedule exports are valid HTML files.')
