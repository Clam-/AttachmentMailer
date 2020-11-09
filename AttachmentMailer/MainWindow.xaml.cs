using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Publisher = Microsoft.Office.Interop.Publisher;

namespace AttachmentMailer
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	/// 

	class DataException : Exception
	{
		public DataException(string message) : base(message) { }
		public DataException(string message, System.Exception inner) : base(message, inner) { }
	}

	internal class Logger
	{
		private static TraceSource logging = new TraceSource("AttachmentMailer");
		internal static void log(TraceEventType tt, int code, string s)
		{
			logging.TraceEvent(tt, code, "\r\n\t" + DateTime.Now.ToString() + "\r\n" + s + "\r\n");
		}
	}

	public partial class MainWindow : Window
	{

		private static string MERGELOC = "merged";
		private static int HASHFIELDNUMS = 15;
		Excel.Application exApp;
		Word.Application wordApp;
		Publisher.Application pubApp;

		Outlook.Application outApp;
		Outlook.NameSpace outNS;

		Outlook.MAPIFolder folderMAPI;

		BackgroundWorker worker = null;

		static string tempDirectory = null;
		static string tempMerge = null;

		Dictionary<string, List<string[]>> mergedocs = null;

		public MainWindow()
		{
			Resources["Attachments"] = new ObservableCollection<Attachment>();
			Resources["Replacements"] = new ObservableCollection<Replacement>();
			Resources["Docs"] = new ObservableCollection<Document>();
			mergedocs = new Dictionary<string, List<string[]>>();
			InitializeComponent();
		}

		private void createDraftsButton_Click(object sender, RoutedEventArgs e)
		{
			Logger.log(TraceEventType.Verbose, 3, "Clicked create drafts.");
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
				statusLabel.Text = "Bad selection.";
				return;
			}

			// check for a draft
			if (folderMAPI != null && (folderMAPI.Items.Count < 1))
			{
				statusLabel.Text = "Cannot find draft email.";
				return;
			}

			// do worker stuff here
			object[] arg = { 
							   ((ObservableCollection<Attachment>)Resources["Attachments"]).ToList(), 
							   ((ObservableCollection<Replacement>)Resources["Replacements"]).ToList(),
							   ((ObservableCollection<Document>)Resources["Docs"]).ToList()
						   };

			if (worker != null && worker.WorkerSupportsCancellation) { worker.CancelAsync(); }
			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;
			worker.DoWork +=
				new DoWorkEventHandler(createDraftsWorker);
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(workerDone);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync(arg);
		}

		private void disableUI()
		{
			inputSections.IsEnabled = false;
			processMergeButton.IsEnabled = false;
			processPublisherButton.IsEnabled = false;
			createDraftsButton.IsEnabled = false;
			sendDraftsButton.IsEnabled = false;
			cancelButton.IsEnabled = true;
		}

		private void createDraftsWorker(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			Logger.log(TraceEventType.Verbose, 3, "Starting Create Drafts worker...");
			if (!checkApps())
			{
				return;
			}
			// check email field
			int emailfield = -1;
			Dispatcher.Invoke(new Action(() =>
			{
				emailfield = getEmailColumn();
			}));

			if (emailfield == -1)
			{
				return;
			}
			Excel.Range selection = exApp.Selection;
			if (selection == null)
			{
				throw new DataException("Bad selection.");
			}
			// get the draft
			Outlook._MailItem orig;
			if (folderMAPI != null && (folderMAPI.Items.Count > 0))
			{
				orig = folderMAPI.Items[1];
			}
			else
			{
				throw new DataException("Cannot find draft email.");
			}
			Outlook.MAPIFolder drafts = outNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
			// unpack
			object[] args = (object[])e.Argument;
			List<Attachment> ds = (List<Attachment>)args[0];
			List<Replacement> drs = (List<Replacement>)args[1];
			List<Document> docs = (List<Document>)args[2];
			if ((docs.Count > 0) && (mergedocs.Count == 0))
			{
				// bail
				throw new DataException("Error: Haven't processed merge it seems.");
			}
			// Let's iterating
			Boolean missingAttachment = false;
			String missingAttachments = null;
			int count = 0;
			int max = selection.Rows.Count;
			Dictionary<String, Outlook._MailItem> mailitems = new Dictionary<string,Outlook._MailItem>();

			foreach (Excel.Range row in selection.Rows)
			{
				String hash = "";
				
				//generate hash
				StringBuilder sb = new StringBuilder();

				int[] ia = Option.getColumns();
				if (ia != null)
				{
					foreach (int xi in ia)
					{
						try { sb.Append(processFloat(getCellContent(row.Cells[xi]))); }
						catch (COMException) { continue; }
						catch (Exception exc) { Logger.log(TraceEventType.Error, 9, exc.ToString() + "\r\nxi:" + xi + "\r\n"); continue; }
					}
				}
				else
				{
					for (int xi = 1; xi <= HASHFIELDNUMS; xi++)
					{
						try { sb.Append(processFloat(getCellContent(row.Cells[xi]))); }
						catch (COMException) { continue; }
						catch (Exception exc) { Logger.log(TraceEventType.Error, 9, exc.ToString() + "\r\nxi:" + xi + "\r\n"); continue; }
					}
				}

				SHA1 sha = new SHA1CryptoServiceProvider();
				hash = BitConverter.ToString(sha.ComputeHash(
						Encoding.Unicode.GetBytes(sb.ToString())
					)).Replace("-", string.Empty);
				Logger.log(TraceEventType.Verbose, 1, "hash:" + hash + " hashed data: " + sb.ToString()); // .Substring(0, 20)
				
				if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
				Outlook._MailItem newMI;
				bool newItem = true;
				if (!Option.createforuniquehash || (Option.createforuniquehash && !mailitems.ContainsKey(hash)))
				{
					Logger.log(TraceEventType.Verbose, 999, "Making new mail for:" + hash);
					try
					{
						newMI = orig.Copy();
						if (Option.createforuniquehash) { mailitems[hash] = newMI; }
					}
					catch (COMException ex)
					{
						Logger.log(TraceEventType.Error, 9, "Outlook Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
						throw new DataException("Cannot open mail in \"Inline view.\" Either browse to a new folder/location in Outlook or disable \"Inline view.\"");
					}
					Logger.log(TraceEventType.Verbose, 1, "Mail item:" + newMI + "||" + newMI.EntryID);
				}
				else {
					newMI = mailitems[hash]; newItem = false;
					Logger.log(TraceEventType.Verbose, 1, "\r\nGETTING MAIL ITEM FOR HASH: " + hash +
						"\r\nMail item:" + newMI + "||" + newMI.EntryID);
				}
				
				string email = getCellContent(row.Cells[emailfield]);
				if (!newItem)
					Logger.log(TraceEventType.Verbose, 1, "\r\nUpdating email ... " + email);
				if (!email.Equals("")){ newMI.To = email; }
				// add existing items
				if (!newItem) { Logger.log(TraceEventType.Verbose, 1, "\r\nDoing attachments for existing mail...");  }
				else { Logger.log(TraceEventType.Verbose, 1, "\r\nDoing attachments for new mail..."); }
				foreach (Attachment d in ds)
				{
					String fname = processAttachmentName(d.attachmentName, row);
                    String fn = fname;
					fname = System.IO.Path.Combine(d.location, fname);
					Logger.log(TraceEventType.Verbose, 1, "\r\nAttaching... " + fname);
					try
					{
                        bool found = false;
                        foreach (Outlook.Attachment a in newMI.Attachments)
                        {
                            if (a.FileName == fn) { found = true; }
                        }
                        if (!found) { newMI.Attachments.Add(fname); }
					}
					catch (FileNotFoundException)
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
				// add merge items. only if newitem:
				if (newItem)
				{
					if (mergedocs.Count > 0 && mergedocs.ContainsKey(hash))
					{
						foreach (string[] fnames in mergedocs[hash])
						{
							string nf = Path.Combine(tempMerge, fnames[1]);
							File.Move(fnames[0], nf);
							Logger.log(TraceEventType.Verbose, 1, "\r\nAttaching... " + nf);
							try
							{
								newMI.Attachments.Add(nf);
							}
							catch (FileNotFoundException)
							{
								if (!missingAttachment)
								{
									missingAttachments = "Attachments missing: (HASH)" + fnames[1];
									missingAttachment = true;
								}
								else
								{
									missingAttachments = String.Concat(missingAttachments, "\n(HASH)" + fnames[1]);
								}
							}
							File.Move(nf, fnames[0]);
						}
					}
					else if (mergedocs.Count > 0)
					{
						if (!missingAttachment)
						{
							missingAttachments = "Attachments missing: (HASH)" + hash;
							missingAttachment = true;
						}
						else
						{
							missingAttachments = String.Concat(missingAttachments, "\n(HASH)" + hash);
						}
					}
				}
				if (drs.Count > 0 && newItem)
				{
					// do body (and Subject) replacements only if newitem
					try
					{
						string msg = newMI.HTMLBody;
						string subject = newMI.Subject;
						foreach (Replacement dr in drs)
						{
							msg = msg.Replace(dr.placeholder, getCellContent(row.Cells[dr.replacement]));
                            try
                            {
                                subject = subject.Replace(dr.placeholder, getCellContent(row.Cells[dr.replacement]));
                            } catch (NullReferenceException)
                            {
                                newMI.Close(Outlook.OlInspectorClose.olDiscard);
                                Marshal.FinalReleaseComObject(newMI);
                                if (Option.createforuniquehash) { cleanUpDraftDict(mailitems, null); }
                                Marshal.FinalReleaseComObject(selection);
                                Marshal.FinalReleaseComObject(drafts);
                                throw new DataException("Missing Subject from Draft");
                            }
						}
						newMI.HTMLBody = msg;
						newMI.Subject = subject;
					}
					catch (COMException ex)
					{
						newMI.Close(Outlook.OlInspectorClose.olDiscard);
						Marshal.FinalReleaseComObject(newMI);
						if (Option.createforuniquehash) { cleanUpDraftDict(mailitems, null); }
						Marshal.FinalReleaseComObject(selection);
						Marshal.FinalReleaseComObject(drafts);
						Logger.log(TraceEventType.Error, 9, "Outlook Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
						throw new DataException("Access denied. You need to change Outlook settings.");
					}
				}
				if (!Option.createforuniquehash) 
				{
					newMI.Move(drafts);
					Marshal.FinalReleaseComObject(newMI); 
				}
				//newMI.Close(Outlook.OlInspectorClose.olSave);
				//newMI.Save();
				count = count + 1;
				worker.ReportProgress((int)(((float)count / (float)max) * 100));
			}
			// clean up dictionary of mailitems
			if (Option.createforuniquehash) { cleanUpDraftDict(mailitems, drafts); }
			Marshal.FinalReleaseComObject(selection);
			Marshal.FinalReleaseComObject(drafts);

			orig.Close(Outlook.OlInspectorClose.olDiscard); //discard
			Marshal.FinalReleaseComObject(orig);
			if (missingAttachment)
			{
				Logger.log(TraceEventType.Information, 9, "Missing: " + missingAttachments);
				e.Result = "Done, but: " + missingAttachments;
			}
			else
			{
				e.Result = "Done.";
			}
			Logger.log(TraceEventType.Verbose, 1, "Create drafts worker done.");
			Logger.log(TraceEventType.Information, 3, "Created " + count + " drafts, with " + docs.Count + " merged docs, " +
				ds.Count + " attachments and " + drs.Count + " replacements.");
		}

		private void cleanUpDraftDict(Dictionary<String, Outlook._MailItem> items, Outlook.MAPIFolder drafts)
		{
			// clean up dictionary of mailitem
			Logger.log(TraceEventType.Verbose, 1, "\r\nCleaning up unique mail dict");
			List<String> keys = items.Keys.ToList();
			foreach (String key in keys)
			{
				Outlook._MailItem mi = items[key];
                if (drafts != null) { mi.Move(drafts); }
                else
                {
                    try { mi.Close(Outlook.OlInspectorClose.olDiscard); }
                    catch (System.Runtime.InteropServices.InvalidComObjectException e) { }
                }
				items.Remove(key);
				Marshal.FinalReleaseComObject(mi);
			}
		}

		private void workerDone(object sender, RunWorkerCompletedEventArgs e)
		{
			// First, handle the case where an exception was thrown. 
			if (e.Error != null)
			{
				statusLabel.Text = e.Error.Message;
				Logger.log(TraceEventType.Error, 9, "Worker Exception\r\n" + e.Error.GetType() + ":" + e.Error.Message + "\r\n" + e.Error.StackTrace);
			}
			else if (e.Cancelled)
			{
				// Next, handle the case where the user canceled  
				// the operation. 
				// Note that due to a race condition in  
				// the DoWork event handler, the Cancelled 
				// flag may not have been set, even though 
				// CancelAsync was called.
				statusLabel.Text = "Operation canceled";
			}
			else
			{
				// Finally, handle the case where the operation  
				// succeeded.
				if (e.Result != null) { statusLabel.Text = e.Result.ToString(); }
				else { statusLabel.Text = "Done."; }
			}

			//set buttons
			inputSections.IsEnabled = true;
			sendDraftsButton.IsEnabled = true;
			createDraftsButton.IsEnabled = true;
			cancelButton.IsEnabled = false;
			processMergeButton.IsEnabled = true;
			processPublisherButton.IsEnabled = true;
		}

		private void addAttachButton_Click(object sender, RoutedEventArgs e)
		{
			if (addAttachFolderLabel.Content.Equals(""))
			{
				statusLabel.Text = "Invalid folder.";
			}
			else if (addAttachFolderLabel.Content.Equals("-1"))
			{
				statusLabel.Text = "Invalid folder.";
			}
			else
			{
				((ObservableCollection<Attachment>)attachmentList.ItemsSource)
					.Add(new Attachment(addAttachFolderLabel.Content.ToString(), addAttachInput.Text));

			}
			addAttachFolderLabel.Content = "";
		}

		private void remAttachButton_Click(object sender, RoutedEventArgs e)
		{
			List<Attachment> dr = new List<Attachment>();
			foreach (Attachment d in attachmentList.SelectedItems)
			{
				dr.Add(d);
			}
			foreach (Attachment d in dr)
			{
				((ObservableCollection<Attachment>)attachmentList.ItemsSource).Remove(d);
			}
		}

		private Attachment getAttachment()
		{
			System.Collections.IEnumerator ie = attachmentList.ItemsSource.GetEnumerator();
			if (!ie.MoveNext())
			{
				return null;
			}
			return ((Attachment)ie.Current);
		}

		private Replacement getReplacement()
		{
			System.Collections.IEnumerator ie = replacementList.ItemsSource.GetEnumerator();
			if (!ie.MoveNext())
			{
				return null;
			}
			return ((Replacement)ie.Current);
		}

		private int getEmailColumn()
		{

			int emailindex = parseNumber(emailColumn.Text);
			if (emailindex == -2)
			{
				statusLabel.Text = "Invalid column number for email address. Must be greater than 0.";
				return -1;
			}
			else if (emailindex == -1)
			{
				statusLabel.Text = "Invalid column number for email address.";
				return -1;
			}
			return emailindex;
		}

		private int parseNumber(string s)
		{
			int num;
			try
			{
				num = Convert.ToInt32(s);
				if (num <= 0) { return -2; }
			}
			catch (FormatException) { return -1; }
			return num;
		}

		private int getReplaceColumn()
		{
			int replaceindex = parseNumber(replaceWithCol.Text);
			if (replaceindex == -2)
			{
				statusLabel.Text = "Invalid column number for replacement. Must be greater than 0.";
				return -1;
			}
			else if (replaceindex == -1)
			{
				statusLabel.Text = "Invalid column number for replacement.";
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
				if (index <= 0)
				{
					statusLabel.Text = "Invalid column number.";
					return null;
				}
				matched = true;
				newaname = newaname.Replace(match.Value, getCellContent(row.Cells[index]));
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

		private Object getFirst(System.Collections.IEnumerator ie)
		{
			if (!ie.MoveNext())
			{
				return null;
			}
			else
			{
				return ie.Current;
			}
		}

		private string getCellContent(Excel.Range cell)
		{
			Object data = cell.Value2;
			if (data != null)
			{
				return data.ToString();
			}
			//Logger.log(TraceEventType.Information, 99, "Cell ("+cell.Address+") contains null\r\n");
			return "";
		}

		private void updateEmailButton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}

			int emailindex = getEmailColumn();

			if ((emailindex == -1))
			{
				return;
			}

			Attachment adata = (Attachment)getFirst(attachmentList.ItemsSource.GetEnumerator());

			Replacement rdata = (Replacement)getFirst(replacementList.ItemsSource.GetEnumerator());

			Excel.Range selection;
			try
			{
				selection = exApp.Selection;
			}
			catch (InvalidCastException ex)
			{
				statusLabel.Text = "Bad selection. Try re-opening Excel.";
				Logger.log(TraceEventType.Error, 9, "Cast Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				return;
			}
			if (selection == null)
			{
				statusLabel.Text = "Bad selection. Try re-opening Excel.";
				return;
			}
			//selection = selection.CurrentRegion;
			foreach (Excel.Range row in selection.Rows)
			{
				previewEmail.Content = getCellContent(row.Cells[emailindex]);
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
					previewReplace.Content = rdata.placeholder.Replace(rdata.placeholder, getCellContent(row.Cells[rdata.replacement]));
				}

				break;
			}
			Marshal.FinalReleaseComObject(selection);
			statusLabel.Text = "Ready.";
		}

		private void draftFolderbutton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			if (folderMAPI != null) { Marshal.FinalReleaseComObject(folderMAPI); }

			statusLabel.Text = "Navigate to draft folder in outlook...";
			//ghetto to make the label update
			Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background,
									  new Action(delegate { }));
			folderMAPI = outNS.PickFolder();
			statusLabel.Text = "Ready.";
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
			else
			{
				statusLabel.Text = "Invalid Outlook draft folder.";
			}
		}

		private bool checkWord()
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			if (wordApp != null) { Marshal.FinalReleaseComObject(wordApp); }
			try
			{
				wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
			}
			catch (COMException ex)
			{
				Logger.log(TraceEventType.Error, 9, "Word Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				statusLabel.Text = "Word couldn't be accessed. Word not open?";
				return false;
			}

			if (wordApp == null)
			{
				statusLabel.Text = "Word couldn't be accessed. Word not open?";
				return false;
			}
			return true;
		}

		private bool checkPub()
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			if (pubApp != null) { Marshal.FinalReleaseComObject(pubApp); }
			try
			{
				pubApp = (Publisher.Application)Marshal.GetActiveObject("Publisher.Application");
			}
			catch (COMException ex)
			{
				Logger.log(TraceEventType.Error, 9, "Publisher Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				statusLabel.Text = "Publisher couldn't be accessed. Publisher not open?";
				return false;
			}

			if (pubApp == null)
			{
				statusLabel.Text = "Publisher couldn't be accessed. Publisher not open?";
				return false;
			}
			return true;
		}

		private bool checkApps(bool excel)
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
			if (excel)
			{
				if (exApp != null) { Marshal.FinalReleaseComObject(exApp); }
				try
				{
					exApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
				}
				catch (COMException ex)
				{
					Logger.log(TraceEventType.Error, 9, "Excel exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
					statusLabel.Text = "Excel couldn't be accessed. Excel not open?";
					return false;
				}

				if (exApp == null)
				{
					statusLabel.Text = "Excel couldn't be accessed. Excel not open?";
					return false;
				}
			}
			if (outApp != null) { Marshal.FinalReleaseComObject(outApp); }
			if (outNS != null) { Marshal.FinalReleaseComObject(outNS); }
			try
			{
				outApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
			}
			catch (COMException ex)
			{
				Logger.log(TraceEventType.Error, 9, "Outlook exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				statusLabel.Text = "Outlook couldn't be accessed. Outlook not open?+";
				return false;
			}
			if (outApp == null)
			{
				statusLabel.Text = "Outlook couldn't be accessed. Outlook not open?";
			}

			outNS = outApp.GetNamespace("MAPI");

			if (outNS == null)
			{
				statusLabel.Text = "Uh oh... Bad things happened.";
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
			Logger.log(TraceEventType.Information, 9, "\r\n\r\nStarting up...");
			tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
			while (Directory.Exists(tempDirectory) || File.Exists(tempDirectory))
			{ tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()); }
			Directory.CreateDirectory(tempDirectory);
			Logger.log(TraceEventType.Information, 9, "Temp dir: " + tempDirectory);
			tempMerge = Path.Combine(tempDirectory, MERGELOC);
			statusLabel.Text = "Ready.";
		}

		private void remReplaceButton_Click(object sender, RoutedEventArgs e)
		{
			List<Replacement> dr = new List<Replacement>();
			foreach (Replacement d in replacementList.SelectedItems)
			{
				dr.Add(d);
			}
			foreach (Replacement d in dr)
			{
				((ObservableCollection<Replacement>)replacementList.ItemsSource).Remove(d);
			}
		}

		private void addReplaceButton_Click(object sender, RoutedEventArgs e)
		{
			int replaceindex = getReplaceColumn();

			if (placeholderText.Text.Equals(""))
			{
				statusLabel.Text = "Invalid placeholder.";
				return;
			}
			else if (replaceindex == -1)
			{
				return;
			}
			else
			{
				((ObservableCollection<Replacement>)replacementList.ItemsSource)
					.Add(new Replacement(placeholderText.Text, replaceindex));

			}
		}

		private void sendDraftsButton_Click(object sender, RoutedEventArgs e)
		{
			Logger.log(TraceEventType.Verbose, 3, "Clicked send drafts.");
			if (!checkApps(false))
			{
				return;
			}

			if (worker != null && worker.WorkerSupportsCancellation) { worker.CancelAsync(); }
			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;
			worker.DoWork +=
				new DoWorkEventHandler(sendDraftsWorker);
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(workerDone);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync();
		}

		private void sendDraftsWorker(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			Outlook.MAPIFolder drafts = outNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
			int count = 0;
			int max = drafts.Items.Count;
			string to = "";
			try
			{
				foreach (Outlook._MailItem mi in drafts.Items)
				{
					if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
					if (mi.To != null && !mi.To.Equals("")) {
						to = mi.To;
						Logger.log(TraceEventType.Information, 3, string.Format("Sending to: {0}", to));
						mi.Send(); count = count + 1;
					}
					worker.ReportProgress((int)(((float)count / (float)max) * 100));
				}
			}
			catch (COMException ex)
			{
				Logger.log(TraceEventType.Critical, 9, "Outlook exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				throw new DataException(string.Format("Access denied. Try changing Outlook's security settings. (To: {0}", to));
			}
			catch (Exception ex)
			{
				Logger.log(TraceEventType.Critical, 9, "\r\nCRASH\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				throw new DataException("Other error. (" + ex.GetType() + "):" + ex.Message);
			}
			Logger.log(TraceEventType.Verbose, 3, "Done send drafts. Attempted to send " + count + " drafts.");
		}

		private void cancelButton_Click(object sender, RoutedEventArgs e)
		{
			Logger.log(TraceEventType.Verbose, 3, "Clicked Cancel.");
			if (worker != null && worker.WorkerSupportsCancellation)
			{
				worker.CancelAsync();
			}
		}

		private void addDocumentButton_Click(object sender, RoutedEventArgs e)
		{
			if (addDocLocationLabel.Content == null || addDocLocationLabel.Content.Equals(""))
			{
				statusLabel.Text = "Document missing. Click Browse to locate merge document.";
				return;
			}

			if (!addDocName.Text.EndsWith(".pdf") && !addDocName.Text.EndsWith(".docx"))
			{
				statusLabel.Text = "Attachment name must end with .pdf or .docx";
				return;
			}
			else if (addDocName.Text.Equals(""))
			{
				statusLabel.Text = "Error: Blank attachment name.";
				return;
			}
			// check if attachment name already exist in list
			ObservableCollection<Document> dL = ((ObservableCollection<Document>)documentList.ItemsSource);
			foreach (Document d in dL)
			{
				if (d.attachmentFormat.Equals(addDocName.Text))
				{
					statusLabel.Text = "Error: Attachment name exists already.";
					return;
				}
			}
			((ObservableCollection<Document>)documentList.ItemsSource)
					.Add(new Document((string)addDocLocationLabel.Content, addDocName.Text));
			nukeTempMerges();
			statusLabel.Text = "Ready.";

		}

		private void addDocFileButton_Click(object sender, RoutedEventArgs e)
		{

			System.Windows.Forms.OpenFileDialog fileBrowser = new System.Windows.Forms.OpenFileDialog();
			fileBrowser.Filter = "Documents (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*";

			if (fileBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				Stream docStream = null;
				try
				{
					if ((docStream = fileBrowser.OpenFile()) != null)
					{
						using (docStream) { } // don't use the stream
					}
				}
				catch (SystemException ex)
				{
					addDocLocationLabel.Content = "";
					Logger.log(TraceEventType.Critical, 9, "File exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
					statusLabel.Text = "Error: Could not read file. Original error: " + ex.Message;
					return;
				}

				// check if have merge doc
				if (!checkWord())
				{
					addDocLocationLabel.Content = "";
					return;
				}
				statusLabel.Text = "Click Yes in word...";
				//ghetto to make the label update
				Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background,
											new Action(delegate { }));

				Word.Document doc = wordApp.Application.Documents.Open(fileBrowser.FileName, ReadOnly: true, Visible: false);
				//Console.WriteLine(doc.MailMerge.State);
				//Console.WriteLine(doc.MailMerge.DataSource.ConnectString);
				//Console.WriteLine(doc.MailMerge.DataSource.QueryString);
				if (doc.MailMerge.State == Word.WdMailMergeState.wdMainAndDataSource)
				{
					addDocLocationLabel.Content = fileBrowser.FileName;
					statusLabel.Text = "Ready.";
					//Console.WriteLine("" + doc.MailMerge.DataSource.DataFields[1].Value + doc.MailMerge.DataSource.DataFields[2].Value +
					//doc.MailMerge.DataSource.DataFields[3].Value);
				}
				else
				{
					addDocLocationLabel.Content = "";
					statusLabel.Text = "Selected document does not have merge data.";
				}
				((Word._Document)doc).Close(SaveChanges: false);
				Marshal.FinalReleaseComObject(doc);
			}
			else
			{
				addDocLocationLabel.Content = "";
			}
		}

		private void remDocButton_Click(object sender, RoutedEventArgs e)
		{
			List<Document> doclist = new List<Document>();
			foreach (Document d in documentList.SelectedItems)
			{
				doclist.Add(d);
			}
			foreach (Document d in doclist)
			{
				((ObservableCollection<Document>)documentList.ItemsSource).Remove(d);
			}
			nukeTempMerges();
		}

		private void Window_Closing(object sender, CancelEventArgs e)
		{
			if (worker != null && worker.WorkerSupportsCancellation)
			{
				worker.CancelAsync();
			}
			GC.Collect();
			GC.WaitForPendingFinalizers();
			if (exApp != null)
			{
				//exApp.Quit();
				Marshal.FinalReleaseComObject(exApp);
			}
			if (wordApp != null)
			{
				//wordApp.Quit();
				Marshal.FinalReleaseComObject(wordApp);
			}
			if (pubApp != null)
			{
				//wordApp.Quit();
				Marshal.FinalReleaseComObject(pubApp);
			}
			if (outApp != null)
			{
				Marshal.FinalReleaseComObject(outApp);
			}
		}

		private void processMergeButton_Click(object sender, RoutedEventArgs e)
		{
			Logger.log(TraceEventType.Verbose, 3, "Clicked process merge.");
			// do worker stuff here
			object arg = ((ObservableCollection<Document>)Resources["Docs"]).ToList();

			if (worker != null && worker.WorkerSupportsCancellation) { worker.CancelAsync(); }
			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;
			worker.DoWork +=
				new DoWorkEventHandler(processMergeWorker);
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(workerDone);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync(arg);
		}

		void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			if (e.UserState != null)
				statusLabel.Text = e.UserState as String;
			progress.Value = e.ProgressPercentage;

		}

		private void nukeTempMerges()
		{
			mergedocs.Clear();
			try { Directory.Delete(tempMerge, true); }
			catch (System.IO.DirectoryNotFoundException) { }
			Directory.CreateDirectory(tempMerge);
		}

		private void printMergeDict()
		{
			Logger.log(TraceEventType.Verbose, 999, "MERGE DICT:");
			StringBuilder sb = new StringBuilder();
			foreach (String key in mergedocs.Keys)
			{
				foreach (string[] fnames in mergedocs[key])
				{
					sb.Append(key + ":" + fnames[0] + "-" + fnames[1]+"\n");
				}
			}
			Logger.log(TraceEventType.Verbose, 999, sb.ToString());
		}

		private void processMergeWorker(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			nukeTempMerges();

			List<Document> docs = (List<Document>)e.Argument;
			foreach (Document d in docs)
			{
				if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
				//process docs

				worker.ReportProgress(0, "Press Yes in Word.");
				Word._Document doc = wordApp.Application.Documents.Open(d.location, ReadOnly: true, Visible: false);
				worker.ReportProgress(0, "Processing...");
				if (doc.MailMerge.State == Word.WdMailMergeState.wdMainAndDataSource)
				{
					doc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
					doc.MailMerge.SuppressBlankLines = true;

					doc.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdLastRecord;

					int maxRec = (int)doc.MailMerge.DataSource.ActiveRecord;
					doc.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdFirstRecord;
					int index = (int)doc.MailMerge.DataSource.ActiveRecord;

					OleDbConnection conn = new OleDbConnection(doc.MailMerge.DataSource.ConnectString.Replace("HDR=YES", "HDR=NO").Replace("HDR=Yes", "HDR=NO").Replace("HDR=yes", "HDR=NO"));
					Logger.log(TraceEventType.Verbose, 900, "Q: " + doc.MailMerge.DataSource.QueryString);
					OleDbCommand command = new OleDbCommand(doc.MailMerge.DataSource.QueryString, conn);
					OleDbDataAdapter adapter = new OleDbDataAdapter(command);
					try
					{
						conn.Open();
					}
					catch (InvalidOperationException olex)
					{
						doc.Close(SaveChanges: false);
						throw olex;
					}
					DataSet data = new DataSet();
					adapter.Fill(data, "datas");
					conn.Close();
					DataTable dt = data.Tables["datas"];

					DataRow headrow = dt.Rows[0];
					Dictionary<String, int> headers = new Dictionary<string, int>();
					for (int x = 0; x < 100; x++)
					{
						try
						{
							string col = headrow[x].ToString().Trim().ToLower();
							if (col.Equals("")) { continue; }
							try { headers.Add(col, x); }
							catch (ArgumentException) { continue; }
						}
						catch (IndexOutOfRangeException) { break; }
					}

					int prev = index;
					bool done = false;
					while (!done)
					{
						if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
						worker.ReportProgress((int)(((float)index / (float)maxRec) * 100));
						Logger.log(TraceEventType.Verbose, 9, "Doc: " + d.location + " (rec: " + index + ")");
						doc.MailMerge.DataSource.FirstRecord = index;
						doc.MailMerge.DataSource.LastRecord = index;

						DataRow olerow = dt.Rows[index];
						// hash field data
						StringBuilder sb = new StringBuilder();
						int[] ia = Option.getColumns();
						if (ia != null)
						{
							foreach (int xi in ia)
							{
								try { sb.Append(processFloat(olerow[xi - 1].ToString())); }
								catch (IndexOutOfRangeException) { continue; }
							}
						}
						else
						{
							for (int xi = 1; xi <= HASHFIELDNUMS; xi++)
							{
								try { sb.Append(processFloat(olerow[xi - 1].ToString())); }
								catch (IndexOutOfRangeException) { continue; }
							}
						}
						SHA1 sha = new SHA1CryptoServiceProvider();
						string hash = BitConverter.ToString(sha.ComputeHash(
								Encoding.Unicode.GetBytes(sb.ToString())
							)).Replace("-", string.Empty);
						Logger.log(TraceEventType.Verbose, 1, "hash:" + hash + " hashed data: " + sb.ToString()); //.Substring(0, 20)
						string attachname = processDocAttachmentName(d.attachmentFormat, olerow, headers);
						string docname = Path.Combine(tempMerge, hash + "-" + attachname);
						if (!mergedocs.ContainsKey(hash))
						{
							mergedocs[hash] = new List<string[]>();
						}

						if (File.Exists(docname) && !Option.allowduplicatehash)
						{
							// bail and show error
							Logger.log(TraceEventType.Verbose, 99, "Non unique: " + docname + "\r\n" + "hash:" + hash + " hashed data: " + sb.ToString());
							doc.Close(SaveChanges: false);
							nukeTempMerges();
							throw new DataException("IMPORTANT ERROR: Data source has non unique data.");
						}
						else if (!File.Exists(docname))
						{
							//d.attachmentName = docname;
							doc.MailMerge.Execute(Pause: false);
							Word._Document nd = wordApp.ActiveDocument;
							nd.ExportAsFixedFormat(OutputFileName: docname,
								ExportFormat: Word.WdExportFormat.wdExportFormatPDF);
							nd.Close(SaveChanges: false);
							mergedocs[hash].Add(new string[] { docname, attachname });
						}

						int skips = 0;
						do
						{
							try { doc.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdNextRecord; }
							catch (COMException)
							{
								done = true;
								break;
							}
							index = (int)doc.MailMerge.DataSource.ActiveRecord;
							// skip previous
							Logger.log(TraceEventType.Verbose, 99, "Skipping (" + prev + "->" + index + ") skipped:" + skips);
							skips = skips + 1;
						} while (index == prev && skips < 10);
						if (prev == index) break;
						prev = index;
					}
				}
				else
				{
					doc.Close(SaveChanges: false);
					throw new DataException("Document does not have merged data: " + d.location);
				}
				doc.Close(SaveChanges: false);
			}
		}

		private void helpButton_Click(object sender, RoutedEventArgs e)
		{
			ProcessStartInfo psi = new ProcessStartInfo("https://github.com/Clam-/AttachmentMailer/wiki/Help");
			Process.Start(psi);
		}

		private void helpButton_Click_test(object sender, RoutedEventArgs e)
		{

		}

		private void Window_Closed(object sender, EventArgs e)
		{
			// clean up temp folder
			if (Directory.Exists(tempDirectory))
			{
				try { Directory.Delete(tempDirectory, true); }
				catch (IOException) { }
			}
		}

		private String processDocAttachmentName(String s, DataRow row, Dictionary<String, int> headers)
		{
			String newaname = String.Copy(s);
			string pattern = @"({[a-zA-Z _.,;:'""-]+})";
			Boolean matched = false;
			foreach (Match match in Regex.Matches(s, pattern))
			{
				string field = match.Value.Trim(new Char[] { '{', '}' }).ToLower();
				if (field.Equals(""))
				{
					statusLabel.Text = "Invalid column name.";
					return null;
				}
				matched = true;
				if (headers.ContainsKey(field))
				{
					newaname = newaname.Replace(match.Value, row[headers[field]].ToString());
				}
				else
				{
					throw new DataException("Data source does not have header labelled (" + field + ")");
				}

			}
			if (matched)
				return newaname;
			else
				return s;
		}

		private void openMergeButton_Click(object sender, RoutedEventArgs e)
		{
			if (Directory.Exists(tempMerge))
			{
				ProcessStartInfo psi = new ProcessStartInfo(tempMerge);
				Process.Start(psi);
			}
			else { statusLabel.Text = "Merges not created yet."; }
		}

		private String processFloat(String s)
		{
			string pattern = @"^[0-9]+\.[0-9]+$";
			if (Regex.IsMatch(s, pattern))
			{
				// truncate float
				Logger.log(TraceEventType.Verbose, 99, "ISFLOAT: " + s);
				s = s.Substring(0, 11);
			}
			return s;
		}

		private void processPublisherButton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkPub()) { return; }
			Logger.log(TraceEventType.Verbose, 3, "Clicked process publisher.");
			// do worker stuff here
			object arg = addDocName.Text;

			if (worker != null && worker.WorkerSupportsCancellation) { worker.CancelAsync(); }
			worker = new BackgroundWorker();
			worker.WorkerSupportsCancellation = true;
			worker.WorkerReportsProgress = true;
			worker.DoWork +=
				new DoWorkEventHandler(processPublisherWorker);
			worker.RunWorkerCompleted +=
				new RunWorkerCompletedEventHandler(workerDone);
			worker.ProgressChanged += worker_ProgressChanged;
			disableUI();
			worker.RunWorkerAsync(arg);
		}

		private void processPublisherWorker(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			Directory.CreateDirectory(tempMerge);

			Publisher._Document doc = pubApp.Application.ActiveDocument;
			if (doc == null) { throw new DataException("No active publisher document open to merge."); }

			string attachFormat = (String)e.Argument;

			if (attachFormat.Equals("")) { throw new DataException("Require attachment name to be set in Merged Documents section."); }

			if ((worker.CancellationPending == true)) { e.Cancel = true; return; }
			//process docs

			worker.ReportProgress(0, "Processing...");
			doc.MailMerge.SuppressBlankLines = false;

			int maxRec = (int)doc.MailMerge.DataSource.RecordCount;
			//doc.MailMerge.DataSource.ActiveRecord = doc.MailMerge.DataSource.FirstRecord;
			int origactive = (int)doc.MailMerge.DataSource.ActiveRecord;
			int index = (int)doc.MailMerge.DataSource.ActiveRecord;

			int pages = doc.Pages.Count;
			int startpage;
			int endpage;

			OleDbConnection conn = new OleDbConnection(doc.MailMerge.DataSource.ConnectString.Replace("HDR=YES", "HDR=NO").
				Replace("HDR=Yes", "HDR=NO").Replace("HDR=yes", "HDR=NO"));
			OleDbCommand command = new OleDbCommand("SELECT * FROM [" + doc.MailMerge.DataSource.TableName + "]", conn);
			OleDbDataAdapter adapter = new OleDbDataAdapter(command);
			
			conn.Open();
			//catch (InvalidOperationException olex) { throw olex; }

			DataSet data = new DataSet();
			adapter.Fill(data, "datas");
			conn.Close();
			DataTable dt = data.Tables["datas"];

			DataRow headrow = dt.Rows[0];
			Dictionary<String, int> headers = new Dictionary<string, int>();
			for (int x = 0; x < 100; x++)
			{
				try
				{
					string col = headrow[x].ToString().Trim().ToLower();
					if (col.Equals("")) { continue; }
					try { headers.Add(col, x); }
					catch (ArgumentException) { continue; }
				}
				catch (IndexOutOfRangeException) { break; }
			}
			DataRow olerow;
			TimeSpan t = (DateTime.UtcNow - new DateTime(1970, 1, 1));
			Logger.log(TraceEventType.Information, 5, "START: " + (int)t.TotalSeconds);

			Publisher._Document nd = doc.MailMerge.Execute(Pause: false, Destination: Publisher.PbMailMergeDestination.pbMergeToNewPublication);
			doc.MailMerge.DataSource.ActiveRecord = origactive;
			while (index <= maxRec)
			{
				if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
				worker.ReportProgress((int)(((float)index / (float)maxRec) * 100));
				Logger.log(TraceEventType.Verbose, 9, "PUB - " + " (rec: " + index + ")");

				try { olerow = dt.Rows[index]; }
				catch (IndexOutOfRangeException) { break; }
				// hash field data
				StringBuilder sb = new StringBuilder();
				int[] ia = Option.getColumns();
				if (ia != null)
				{
					foreach (int xi in ia)
					{
						try { sb.Append(processFloat(olerow[xi - 1].ToString())); }
						catch (IndexOutOfRangeException) { continue; }
					}
				}
				else
				{
					for (int xi = 1; xi <= HASHFIELDNUMS; xi++)
					{
						try { sb.Append(processFloat(olerow[xi - 1].ToString())); }
						catch (IndexOutOfRangeException) { continue; }
					}
				}
				SHA1 sha = new SHA1CryptoServiceProvider();
				string hash = BitConverter.ToString(sha.ComputeHash(
						Encoding.Unicode.GetBytes(sb.ToString())
					)).Replace("-", string.Empty);
				Logger.log(TraceEventType.Verbose, 1, "hash:" + hash + " hashed data: " + sb.ToString()); //.Substring(0, 20)
				string attachname = processDocAttachmentName(attachFormat, olerow, headers);
				string docname = Path.Combine(tempMerge, hash + "-" + attachname);
				if (!mergedocs.ContainsKey(hash))
				{
					mergedocs[hash] = new List<string[]>();
				}

				if (File.Exists(docname) && !Option.allowduplicatehash)
				{
					// bail and show error
					Logger.log(TraceEventType.Verbose, 99, "Non unique: " + docname + "\r\n" + "hash:" + hash + " hashed data: " + sb.ToString());
					nukeTempMerges();
					throw new DataException("IMPORTANT ERROR: Data source has non unique data. Or same attachment name.");
				}
				else if (!File.Exists(docname))
				{
					Logger.log(TraceEventType.Verbose, 90, "Saving to: " + docname);

					startpage = index * pages;
					endpage = startpage + pages - 1;
					nd.ExportAsFixedFormat(Publisher.PbFixedFormatType.pbFixedFormatTypePDF, docname, From: startpage, To: endpage);
					mergedocs[hash].Add(new string[] { docname, attachname });
				}
				index = index + 1;
				Logger.log(TraceEventType.Verbose, 99, "Trying: " + index);
			}
			nd.Close();
			t = (DateTime.UtcNow - new DateTime(1970, 1, 1));
			Logger.log(TraceEventType.Information, 5, "END: " + (int)t.TotalSeconds);
		}

		private void advancedButton_Click(object sender, RoutedEventArgs e)
		{
			var options = new Advanced();
			options.Owner = this;
			options.ShowDialog();
		}

	}
}
