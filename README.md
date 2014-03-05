Attachment Mailer for .net (Uses Microsoft Excel and Microsoft Outlook.

This is a program that will help with sending attachments with filenames that are generated from a spreadsheet to email addreses in that same sheet.

It currently uses a draft e-mail you have created (in a separate Outlook folder) and generates the attachment mails from that in your Draft folder. (Tip: You can press Enter to open a message then ALT+S to send. Repeat these 2 key presses to send many e-mails easily. This program will eventually do the sending for you but not at the moment.)

Common workflow is:

1. Select the cells that contain the email addresses and the paramters for the filenames in Excel (leave Excel open.)
2. Enter the column number of the column that contains the email addresses.
3. Type the filename format of the attachment filenames. E.g:
	* Original Filename: "_Kristy Summers appointment document.pdf_"
	* Filename in program: "_{3} {2} appointment document.pdf_"
	* Where 3 and 2 are the column numbers for the first name and last name.
4. Click Browse to select the location of the attachment.
5. After selecting the location click Add
6. Click Update Preview to make sure that it looks right:
	* Preview displays _\<destination email address\>_ **->** _\<first generated attachment filename\>_
7. Press "_Browse Outlook Folders..._" to select the folder that the draft email resides in.
	* **Note:** The draft e-mail should be the only item in the folder. The program uses the "first" item in the folder for simplicity.
7. Press Go
8. Draft e-mails should now sit in your Drafts folder ready to be sent.


**CAVEATS**  
There is not much error checking in the program at the moment. The program will list any attachments that were not found during processing after completion, and warn you if Excel or Outlook isn't open. But that's about it.

**TODOs**
Implement some auto-PDFing/merging functionality so attachments don't need to be generated first.  
Make the program a bit more user friendly.  
etc...  