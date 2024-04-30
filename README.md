# Payroll-Reformatting-Automation
This excel worksheet has a VBA macro that takes an input spreadsheet direct from a specific payroll company's reports and allows the accountants to reconcile it easily. This was made for a specific company's internal process. All names and other data have been changed for privacy purposes.

To use the spreadsheet, download it. After opening it, hit the button on the first sheet. A dialog box will open asking if you want to select an outside file or use the "RAW" sheet. Select based on your preference. Once the process is done, go to the "Output" sheet and reconcile the reformatted data. 

Not all input files will have the same columns, since the payroll vendor's output omits columns when unused, so this process is deliberately greedy at the accountant's request - it will add columns to the main formatting sheet but never get rid of them. They will be inserted in the correct order, not added to the end of the sheet. 

The cells highlighted in yellow in the "Output" sheet are entered from a separate report ot make sure this report is accurate to the timesheets submitted.
