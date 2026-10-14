# Faculty loading document generator

## Install and start

Install Python 3.10 or newer, then in this folder run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open the local address displayed in the terminal. Upload the two HTML schedule exports and the current faculty roster (.xlsx). The roster must have Instructor, Loading %, Role, Chair, Admin, VCAA in columns A–F. Only roster faculty are generated; the role comes from the roster. Enter an optional semester label and download individual workbooks/email drafts or a ZIP containing all successful results. Do not upload the Excel template in the app; a copy is bundled as `template.xlsx`.

Cell O19 (Academic Success) is filled as 0 for Deans and Adjuncts, 8 for other roles in Fall/Spring, and 6.25 for other roles in Summer. The program uses `overload_table.xlsx` to fill acceptable credit hours (maximum only) and acceptable contact hours (the listed range). Academic periods ending 10 use summer values; those ending 20 or 30 use fall/spring values. The source does not specify a minimum credit hour load. When either cap is exceeded, it suggests the fewest overload sections needed (one or two), excluding campus P1. Among equally sized valid selections it favors staying above the table’s minimum weekly contact load, then retaining the most combined credit/contact capacity. If no eligible selection of at most two courses works, generation stops with a review message. The suggestion must be checked against actual contracts and assignments.

The program combines both exports, checks that their academic periods match, excludes cancelled and unassigned sections, and collapses repeated meeting rows by CRN and instructor ID. It retains the template's layout and review fields while removing all conditional formatting. The weekly and semester contact totals exclude marked overload sections. The form has separate Clinical, Externships, Simulations, and Academic Success columns in place of Other Contact Hours. Column Q stores Instructor Responsibility (%) as an Excel percentage. Columns I and U calculate source Credits and Contact Hours multiplied by Q. These adjusted values are used for both credit-load and contact-load overload selection. Missing or invalid responsibility percentages stop generation for review; its breakdown among these categories, suggested overload status, contract percentage, and role require review. The course area expands automatically beyond ten assignments, moving the totals and review fields down and extending their formulas. Enrollment and assignments reflect the export date. Email drafts (.eml) contain the workbook attachment, are addressed to the Chair, and CC Admin and VCAA from the roster. X-Unsent marks each message as a draft for clients that support it. Nothing is sent or created in an email account automatically; client behavior varies and some clients open .eml as read-only messages. Review recipients and content in your email client before sending. Approval tracking is not included.

Loading % supports raw numbers such as 100 and 75, percentage-formatted Excel cells, or text such as 75%. It populates Load Request on the total/review section, moving down when course rows expand. It does not scale role limits. Names match case-insensitively with whitespace normalized; missing or ambiguous schedule matches are reported and skipped. Invalid roster rows or documents that cannot meet the overload rules are reported and skipped while other faculty continue.

The included `template.xlsx` is based on the supplied Fall 2025 workbook. Replace it with an updated workbook with a `Loading Template` worksheet if your form changes, and review field locations in `loading.py`.
