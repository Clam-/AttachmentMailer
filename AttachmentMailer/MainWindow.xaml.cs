using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace AttachmentMailer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    // http://stackoverflow.com/a/159419  Keep this in mind. For Excel cleanup.

	public partial class MainWindow : Window
	{

		Excel.Application exApp;

		Outlook.Application outApp;
		Outlook.NameSpace outNS;

		ListBox attachments;
		TextBox attachmentName;
		Label attachmentFolder;
		Label status;

		Outlook.MAPIFolder folderMAPI;

		public MainWindow()
		{
			InitializeComponent();

		}

		private void doThings_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			
			// check email field
			int emailfield = getEmailColumn();
			// check attachments
			getAttachment();

			Excel.Range selection = exApp.Selection;
			if (selection == null)
			{
				status.Content = "Bad selection.";
				return;
			}

			// get the draft
			Outlook._MailItem orig;
			if (folderMAPI != null && (folderMAPI.Items.Count > 0)) {
				orig = folderMAPI.Items[1];
			} else {
				status.Content = "Cannot find draft email.";
				return;
			}

			Outlook.MAPIFolder drafts = outNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);

			// Let's iterating
			Boolean missingAttachment = false;
			String missingAttachments = null;
			foreach (Excel.Range row in selection.Rows)
			{
				Outlook._MailItem newMI = orig.Copy();
				newMI.To = row.Cells[emailfield].Value2;
				foreach (Data d in attachments.ItemsSource) {
					String fname = processAttachmentName(d.attachmentName, row);
					fname = System.IO.Path.Combine(d.location, fname);
					try
					{
						newMI.Attachments.Add(fname);
					}
					catch (System.IO.FileNotFoundException ex)
					{
						if (!missingAttachment)
						{
							missingAttachments = "Attachments missing: " + fname;
							missingAttachment = true;
						}
						else
						{
							missingAttachments = String.Concat(missingAttachments, "\n" + fname);
						}
					}
				}
				newMI.Move(drafts);
				//newMI.Close(Outlook.OlInspectorClose.olSave);
			}

			if (missingAttachment) {
				Console.WriteLine("Missing: " + missingAttachments);
				status.Content = missingAttachments;
			}
			else
			{
				status.Content = "Done.";
			}
			orig.Close(Outlook.OlInspectorClose.olDiscard); //discard

		}

		private void addAttachButton_Click(object sender, RoutedEventArgs e)
		{
			if (attachmentFolder.Content.Equals(""))
			{
				status.Content = "Invalid folder.";
			} else if (attachmentFolder.Content.Equals("-1"))
			{
				status.Content = "Invalid folder.";
			} else
			{
				((Datum)attachments.ItemsSource).Add(new Data(attachmentFolder.Content.ToString(), attachmentName.Text));
				
			}
			attachmentFolder.Content = "";
		}

		private void remAttachButton_Click(object sender, RoutedEventArgs e)
		{
			List<Data> dr = new List<Data>();
			foreach( Data d in attachments.SelectedItems) {
				dr.Add(d);
			}
			foreach (Data d in dr)
			{
				((Datum)attachments.ItemsSource).Remove(d);
			}
		}

		private Data getAttachment()
		{
			System.Collections.IEnumerator ie = attachments.ItemsSource.GetEnumerator();
			if (!ie.MoveNext())
			{
				status.Content = "No attachments.";
				return null;
			}
			return ((Data)ie.Current);
		}

		private int getEmailColumn()
		{
			int emailindex;
			try
			{
				emailindex = Convert.ToInt32(emailColumn.Text);
				if (emailindex <= 0)
				{
					status.Content = "Invalid column number for email address. Must be greater than 0.";
					return -1;
				}
			}
			catch (FormatException)
			{
				status.Content = "Invalid column number for email address.";
				return -1;
			}
			return emailindex;
		}



		private void addFolderButton_Click(object sender, RoutedEventArgs e)
		{
			System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
			folderBrowser.Description = "Select the folder for this attachment.";
			folderBrowser.ShowNewFolderButton = false;
			System.Windows.Forms.DialogResult result = folderBrowser.ShowDialog();
			if (result == System.Windows.Forms.DialogResult.OK)
			{
				attachmentFolder.Content = folderBrowser.SelectedPath;
			}
			else
			{
				attachmentFolder.Content = "-1";
			}
		}

		private String processAttachmentName(String s, Excel.Range row)
		{
			String newaname = String.Copy(s);
			string pattern = @"({[0-9]+})";
			Boolean matched = false;
			foreach (Match match in Regex.Matches(s, pattern))
			{
				int index = Convert.ToInt32(match.Value.Trim(new Char[] { '{', '}' }));
				Console.WriteLine("Match trim: " + index);
				if (index <= 0)
				{
					status.Content = "Invalid column number.";
					return null;
				}
				matched = true;
				newaname = newaname.Replace(match.Value, row.Cells[index].Value2);
			}
			if (matched)
			{
				return newaname;
			}
			else
			{
				return s;
			}
		}

		private void updateEmailButton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			
			Label previewEmail = (Label)this.FindName("previewEmail");
			Label previewAttachment = (Label)this.FindName("previewAttachment");
			TextBox emailColumn = (TextBox) this.FindName("emailColumn");

			int emailindex = getEmailColumn();
			if (emailindex == -1)
			{
				return;
			}

			Data adata = getAttachment();
			if (adata == null) {
				return;
			}

			Excel.Range selection = exApp.Selection;
			if (selection == null)
			{
				status.Content = "Bad selection. Try re-opening Excel.";
				return;
			}
			//selection = selection.CurrentRegion;
			foreach (Excel.Range row in selection.Rows)
			{
				previewEmail.Content = row.Cells[emailindex].Value2;
				previewAttachment.Content = processAttachmentName(adata.attachmentName, row);
				
				break;
			}
			status.Content = "Ready.";
		}

		private void draftFolderbutton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			
			folderMAPI = outNS.PickFolder();

			//OutlookFolderDialog ofd = new OutlookFolderDialog();
			//if (ofd.ShowDialog() == true)
			//{
			//	this.draftFolder.Content = ofd.selectedFolder;

			//}
			//else
			//{
			//	this.draftFolder.Content = "";
			//}
			if (folderMAPI != null)
			{
				this.draftFolder.Content = folderMAPI.FolderPath;
			}
		}

		private bool checkApps()
		{
			try
			{
				exApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				status.Content = "Excel couldn't be accessed. Excel not open?";
				return false;
			}

			if (exApp == null)
			{
				status.Content = "Excel couldn't be accessed. Excel not open?";
				return false;
			}

			try
			{
				outApp = (Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				status.Content = "Outlook couldn't be accessed. Outlook not open?" + ex;
				return false;
			}
			if (outApp == null)
			{
				status.Content = "Outlook couldn't be accessed. Outlook not open?";
			}
			outNS = outApp.GetNamespace("MAPI");

			if (outNS == null)
			{
				status.Content = "Uh oh... Bad things happened.";
				return false;
			}
			return true;
		}

		private void mainListLayout_Loaded(object sender, RoutedEventArgs e)
		{
			attachments = (ListBox)this.FindName("attachmentList");
			attachmentName = (TextBox)this.FindName("addAttachInput");
			attachmentFolder = (Label)this.FindName("addFolderLabel");
			status = (Label)this.FindName("msgLabel");

			status.Content = "Ready.";
			Console.WriteLine("\nREADY TO DO THINGS\n");

		}
	}
}
