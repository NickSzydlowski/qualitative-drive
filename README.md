# Qualitative Coding Tool for Google Drive
Script to extract comments from Google Drive documents in order to facilitate using Google Drive to code interviews using qualitative research methods.

The script is adapted to the following scenario: a group of researchers are performing qualitative coding on a set of interviews or other texts. The texts are stored in a Shared Drive as Google Documents, and researchers assign codes by highlighting text in the documents and adding comments containing the values that they are assigning to the highlighted text. After coding has been performed, all of the coded documents are stored in a single folder on the Shared Drive. The script then extracts the comments and highlighted text from all documents in the folder, along with metadata identifying the file from which the comments originate.

### Try the script
- Add the script Code.GS to a Google Drive spreadsheet using the spreadsheet's script editor
- Add your Shared Drive ID on line 8 and your folder ID on Line 10
- Run the script from the spreadsheet's script editor 
- - The script will populate the spreadsheet with information about the comments for all documents in the specified folder
