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

		Outlook.MAPIFolder folderMAPI;

		public MainWindow()
		{
			Resources["Datum"] = new Datum();
			Resources["DataReplacements"] = new DataReplacements();
			InitializeComponent();
		}

		private void createDraftsButton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			
			// check email field
			int emailfield = getEmailColumn();
			if (emailfield == -1)
			{
				return;
			}
			Excel.Range selection = exApp.Selection;
			if (selection == null)
			{
				statusLabel.Content = "Bad selection.";
				return;
			}

			// get the draft
			Outlook._MailItem orig;
			if (folderMAPI != null && (folderMAPI.Items.Count > 0)) {
				orig = folderMAPI.Items[1];
			} else {
				statusLabel.Content = "Cannot find draft email.";
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
				foreach (Data d in attachmentList.ItemsSource)
				{
					String fname = processAttachmentName(d.attachmentName, row);
					fname = System.IO.Path.Combine(d.location, fname);
					try
					{
						newMI.Attachments.Add(fname);
					}
					catch (System.IO.FileNotFoundException)
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
				try 
				{
					// do body replacements
					try
					{
						string msg = newMI.HTMLBody;
						foreach (DataReplace dr in replacementList.ItemsSource)
						{
							msg = msg.Replace(dr.placeholder, row.Cells[dr.replacement].Value2);
						}
						newMI.HTMLBody = msg;
					}
					catch (System.Runtime.InteropServices.COMException)
					{
						statusLabel.Content = "Access denied. You need to change Outlook settings";
						newMI.Close(Outlook.OlInspectorClose.olDiscard);
						return;
					}
				}
				catch (Exception)
				{
					statusLabel.Content = "Something broke.";
					return;
				}
								

				newMI.Move(drafts);
				//newMI.Close(Outlook.OlInspectorClose.olSave);
			}

			if (missingAttachment) {
				Console.WriteLine("Missing: " + missingAttachments);
				statusLabel.Content = missingAttachments;
			}
			else
			{
				statusLabel.Content = "Done.";
			}
			orig.Close(Outlook.OlInspectorClose.olDiscard); //discard

		}

		private void addAttachButton_Click(object sender, RoutedEventArgs e)
		{
			if (addAttachFolderLabel.Content.Equals(""))
			{
				statusLabel.Content = "Invalid folder.";
			}
			else if (addAttachFolderLabel.Content.Equals("-1"))
			{
				statusLabel.Content = "Invalid folder.";
			} else
			{
				((Datum)attachmentList.ItemsSource).Add(new Data(addAttachFolderLabel.Content.ToString(), addAttachInput.Text));
				
			}
			addAttachFolderLabel.Content = "";
		}

		private void remAttachButton_Click(object sender, RoutedEventArgs e)
		{
			List<Data> dr = new List<Data>();
			foreach (Data d in attachmentList.SelectedItems)
			{
				dr.Add(d);
			}
			foreach (Data d in dr)
			{
				((Datum)attachmentList.ItemsSource).Remove(d);
			}
		}

		private Data getAttachment()
		{
			System.Collections.IEnumerator ie = attachmentList.ItemsSource.GetEnumerator();
			if (!ie.MoveNext())
			{
				return null;
			}
			return ((Data)ie.Current);
		}

		private DataReplace getReplacement()
		{
			System.Collections.IEnumerator ie = replacementList.ItemsSource.GetEnumerator();
			if (!ie.MoveNext())
			{
				return null;
			}
			return ((DataReplace)ie.Current);
		}

		private int getEmailColumn()
		{
			int emailindex;
			try
			{
				emailindex = Convert.ToInt32(emailColumn.Text);
				if (emailindex <= 0)
				{
					statusLabel.Content = "Invalid column number for email address. Must be greater than 0.";
					return -1;
				}
			}
			catch (FormatException)
			{
				statusLabel.Content = "Invalid column number for email address.";
				return -1;
			}
			return emailindex;
		}

		private int getReplaceColumn()
		{
			int replaceindex;
			try
			{
				replaceindex = Convert.ToInt32(replaceWithCol.Text);
				if (replaceindex <= 0)
				{
					statusLabel.Content = "Invalid column number for replacement. Must be greater than 0.";
					return -1;
				}
			}
			catch (FormatException)
			{
				statusLabel.Content = "Invalid column number for replacement.";
				return -1;
			}
			return replaceindex;
		}


		private void addFolderButton_Click(object sender, RoutedEventArgs e)
		{
			System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
			folderBrowser.Description = "Select the folder for this attachment.";
			folderBrowser.ShowNewFolderButton = false;
			System.Windows.Forms.DialogResult result = folderBrowser.ShowDialog();
			if (result == System.Windows.Forms.DialogResult.OK)
			{
				addAttachFolderLabel.Content = folderBrowser.SelectedPath;
			}
			else
			{
				addAttachFolderLabel.Content = "-1";
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
					statusLabel.Content = "Invalid column number.";
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
			int replaceindex = getReplaceColumn();
			if ( (emailindex == -1) || (replaceindex == -1) )
			{
				return;
			}

			Data adata = getAttachment();
			DataReplace rdata = getReplacement();
			
			Excel.Range selection = exApp.Selection;
			if (selection == null)
			{
				statusLabel.Content = "Bad selection. Try re-opening Excel.";
				return;
			}
			//selection = selection.CurrentRegion;
			foreach (Excel.Range row in selection.Rows)
			{
				previewEmail.Content = row.Cells[emailindex].Value2;
				if (adata == null)
				{
					previewAttachment.Content = "";
				}
				else
				{
					previewAttachment.Content = processAttachmentName(adata.attachmentName, row);
				}
				if (rdata == null)
				{
					previewReplace.Content = "";
					previewPlaceholder.Content = "";
				}
				else
				{
					previewPlaceholder.Content = rdata.placeholder;
					previewReplace.Content = rdata.placeholder.Replace(rdata.placeholder, row.Cells[rdata.replacement].Value2);
				}
				
				break;
			}
			statusLabel.Content = "Ready.";
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

		private bool checkApps(bool excel)
		{
			if (excel)
			{
				try
				{
					exApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
				}
				catch (System.Runtime.InteropServices.COMException)
				{
					statusLabel.Content = "Excel couldn't be accessed. Excel not open?";
					return false;
				}

				if (exApp == null)
				{
					statusLabel.Content = "Excel couldn't be accessed. Excel not open?";
					return false;
				}
			}
			try
			{
				outApp = (Outlook.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application");
			}
			catch (System.Runtime.InteropServices.COMException ex)
			{
				statusLabel.Content = "Outlook couldn't be accessed. Outlook not open?" + ex;
				return false;
			}
			if (outApp == null)
			{
				statusLabel.Content = "Outlook couldn't be accessed. Outlook not open?";
			}
			outNS = outApp.GetNamespace("MAPI");

			if (outNS == null)
			{
				statusLabel.Content = "Uh oh... Bad things happened.";
				return false;
			}
			return true;
		}

		private bool checkApps()
		{
			return checkApps(true);
		}

		private void mainListLayout_Loaded(object sender, RoutedEventArgs e)
		{
			statusLabel.Content = "Ready.";
			Console.WriteLine("\nREADY TO DO THINGS\n");

		}

		private void remReplaceButton_Click(object sender, RoutedEventArgs e)
		{
			List<DataReplace> dr = new List<DataReplace>();
			foreach (DataReplace d in replacementList.SelectedItems)
			{
				dr.Add(d);
			}
			foreach (DataReplace d in dr)
			{
				((DataReplacements)replacementList.ItemsSource).Remove(d);
			}
		}

		private void addReplaceButton_Click(object sender, RoutedEventArgs e)
		{
			int replaceindex = getReplaceColumn();

			if (placeholderText.Text.Equals(""))
			{
				statusLabel.Content = "Invalid placeholder.";
				return;
			}
			else if (replaceindex == -1)
			{
				return;
			}
			else
			{
				((DataReplacements)replacementList.ItemsSource).Add(new DataReplace(placeholderText.Text, replaceindex));

			}
		}

		private void sendDraftsButton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps(false))
			{
				return;
			}
			Outlook.MAPIFolder drafts = outNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
			try {
				foreach (Outlook._MailItem mi in drafts.Items)
				{
					mi.Send();
				}
			} catch (System.Runtime.InteropServices.COMException) {
				statusLabel.Content = "Access denied. Try changing Outlook's security settings.";
				return;
			} catch (Exception) {
				statusLabel.Content = "Other error. What happen?";
				return;
			}
		}
	}
}
