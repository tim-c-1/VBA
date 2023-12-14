# Compile Report Save and Send
### A VBA Module for Creating Report Copies from Power Query Data

This VBA module is intended for creating copies of your Power Query report as new files. It saves the imported information in a new workbook,
removes any links to other worksheets, and drafts an email to send the report.


Useful bits of code in this module include:
  - For loop for finding all links in workbook
  - Dynamic file paths
  - Loop through empty columns to clean report
  - Add file hyperlink to email body
