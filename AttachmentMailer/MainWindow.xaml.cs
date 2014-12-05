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

	internal class Logger {
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
				statusLabel.Content = "Bad selection.";
				return;
			}

			// check for a draft
			if (folderMAPI != null && (folderMAPI.Items.Count < 1)) {
				statusLabel.Content = "Cannot find draft email.";
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
			if (folderMAPI != null && (folderMAPI.Items.Count > 0)) {
				orig = folderMAPI.Items[1];
			} else {
				throw new DataException("Cannot find draft email.");
			}
			Outlook.MAPIFolder drafts = outNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts);
			// unpack
			object[] args = (object[]) e.Argument;
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
			foreach (Excel.Range row in selection.Rows)
			{
				if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
				Outlook._MailItem newMI;
				try
				{
					newMI = orig.Copy();
				} catch (COMException ex) {
					Logger.log(TraceEventType.Error, 9, "Outlook Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
					throw new DataException("Cannot open mail in \"Inline view.\" Either browse to a new folder/location in Outlook or disable \"Inline view.\"");
				}
				newMI.To = getCellContent(row.Cells[emailfield]);
				foreach (Attachment d in ds)
				{
					String fname = processAttachmentName(d.attachmentName, row);
					fname = System.IO.Path.Combine(d.location, fname);
					try
					{
						newMI.Attachments.Add(fname);
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
				if (docs.Count > 0)
				{
					//generate hash
					StringBuilder sb = new StringBuilder();
					for (int xi = 1; xi <= HASHFIELDNUMS; xi++)
					{
						try {
							sb.Append(processFloat(getCellContent(row.Cells[xi])));
						}
						catch (COMException) { continue; }
						catch (Exception exc) { Logger.log(TraceEventType.Error, 9, exc.ToString() + "\r\nxi:" + xi + "\r\n"); continue; }
					}
					SHA1 sha = new SHA1CryptoServiceProvider();
					string hash = BitConverter.ToString(sha.ComputeHash(
							Encoding.Unicode.GetBytes(sb.ToString())
						)).Replace("-", string.Empty);
					Logger.log(TraceEventType.Verbose, 1, "hash:" + hash + " hashed data: " + sb.ToString()); // .Substring(0, 20)
					if (mergedocs.ContainsKey(hash))
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
					else
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
				if (drs.Count > 0)
				{
					// do body (and Subject) replacements
					try
					{
						string msg = newMI.HTMLBody;
						string subject = newMI.Subject;
						foreach (Replacement dr in drs)
						{
							msg = msg.Replace(dr.placeholder, getCellContent(row.Cells[dr.replacement]));
							subject = subject.Replace(dr.placeholder, getCellContent(row.Cells[dr.replacement]));
						}
						newMI.HTMLBody = msg;
						newMI.Subject = subject;
					}
					catch (COMException ex)
					{
						newMI.Close(Outlook.OlInspectorClose.olDiscard);
						Marshal.FinalReleaseComObject(newMI);
						Marshal.FinalReleaseComObject(selection);
						Marshal.FinalReleaseComObject(drafts);
						Logger.log(TraceEventType.Error, 9, "Outlook Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
						throw new DataException("Access denied. You need to change Outlook settings.");
					}
				}
				newMI.Move(drafts);
				Marshal.FinalReleaseComObject(newMI);
				//newMI.Close(Outlook.OlInspectorClose.olSave);
				count = count + 1;
				worker.ReportProgress((int)(((float)count / (float)max) * 100));
			}
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

		private void workerDone(object sender, RunWorkerCompletedEventArgs e)
		{
			// First, handle the case where an exception was thrown. 
			if (e.Error != null)
			{
				statusLabel.Content = e.Error.Message;
				Logger.log(TraceEventType.Error, 9,  "Worker Exception\r\n" + e.Error.GetType() + ":" + e.Error.Message + "\r\n" + e.Error.StackTrace);
			}
			else if (e.Cancelled)
			{
				// Next, handle the case where the user canceled  
				// the operation. 
				// Note that due to a race condition in  
				// the DoWork event handler, the Cancelled 
				// flag may not have been set, even though 
				// CancelAsync was called.
				statusLabel.Content = "Operation canceled";
			}
			else
			{
				// Finally, handle the case where the operation  
				// succeeded.
				if (e.Result != null) { statusLabel.Content = e.Result.ToString(); }
				else { statusLabel.Content = "Done.";  }
			}

			//set buttons
			inputSections.IsEnabled = true;
			sendDraftsButton.IsEnabled = true;
			createDraftsButton.IsEnabled = true;
			cancelButton.IsEnabled = false;
			processMergeButton.IsEnabled = true;
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

		private int getEmailColumn() {

			int emailindex = parseNumber(emailColumn.Text);
			if (emailindex == -2)
			{
				statusLabel.Content = "Invalid column number for email address. Must be greater than 0.";
				return -1;
			}
			else if (emailindex == -1)
			{
				statusLabel.Content = "Invalid column number for email address.";
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
				statusLabel.Content = "Invalid column number for replacement. Must be greater than 0.";
				return -1;
			}
			else if (replaceindex == -1)
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
				if (index <= 0)
				{
					statusLabel.Content = "Invalid column number.";
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

			if ( (emailindex == -1) )
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
				statusLabel.Content = "Bad selection. Try re-opening Excel.";
				Logger.log(TraceEventType.Error, 9, "Cast Exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				return;
			}
			if (selection == null)
			{
				statusLabel.Content = "Bad selection. Try re-opening Excel.";
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
			statusLabel.Content = "Ready.";
		}

		private void draftFolderbutton_Click(object sender, RoutedEventArgs e)
		{
			if (!checkApps())
			{
				return;
			}
			if (folderMAPI != null) { Marshal.FinalReleaseComObject(folderMAPI); }

			statusLabel.Content = "Navigate to draft folder in outlook...";
			//ghetto to make the label update
			Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background,
									  new Action(delegate { }));
			folderMAPI = outNS.PickFolder();
			statusLabel.Content = "Ready.";
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
				statusLabel.Content = "Invalid Outlook draft folder.";
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
				statusLabel.Content = "Word couldn't be accessed. Word not open?";
				return false;
			}

			if (wordApp == null)
			{
				statusLabel.Content = "Word couldn't be accessed. Word not open?";
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
					statusLabel.Content = "Excel couldn't be accessed. Excel not open?";
					return false;
				}

				if (exApp == null)
				{
					statusLabel.Content = "Excel couldn't be accessed. Excel not open?";
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
				statusLabel.Content = "Outlook couldn't be accessed. Outlook not open?+";
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
			Logger.log(TraceEventType.Information, 9, "\r\n\r\nStarting up...");
			tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
			while (Directory.Exists(tempDirectory) || File.Exists(tempDirectory))
			{ tempDirectory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()); }
			Directory.CreateDirectory(tempDirectory);
			Logger.log(TraceEventType.Information, 9, "Temp dir: " + tempDirectory);
			tempMerge = Path.Combine(tempDirectory, MERGELOC);
			statusLabel.Content = "Ready.";
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
				statusLabel.Content = "Invalid placeholder.";
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
			try {
				foreach (Outlook._MailItem mi in drafts.Items)
				{
					if ((worker.CancellationPending == true)) { e.Cancel = true; break; }
					if (mi.To != null && !mi.To.Equals("")) { mi.Send(); count = count + 1; }
					worker.ReportProgress((int)(((float)count / (float)max) * 100));
				}
			} catch (COMException ex) {
				Logger.log(TraceEventType.Critical, 9, "Outlook exception\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
				throw new DataException("Access denied. Try changing Outlook's security settings.");
			} catch (Exception ex) {
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
				statusLabel.Content = "Document missing. Click Browse to locate merge document.";
				return;
			}

			if (!addDocName.Text.EndsWith(".pdf") && !addDocName.Text.EndsWith(".docx"))
			{
				statusLabel.Content = "Attachment name must end with .pdf or .docx";
				return;
			}
			else if (addDocName.Text.Equals(""))
			{
				statusLabel.Content = "Error: Blank attachment name.";
				return;
			}
			// check if attachment name already exist in list
			ObservableCollection<Document> dL = ((ObservableCollection<Document>)documentList.ItemsSource);
			foreach (Document d in dL)
			{
				if (d.attachmentFormat.Equals(addDocName.Text))
				{
					statusLabel.Content = "Error: Attachment name exists already.";
					return;
				}
			}
			((ObservableCollection<Document>)documentList.ItemsSource)
					.Add(new Document((string)addDocLocationLabel.Content, addDocName.Text));
			mergedocs.Clear();
			statusLabel.Content = "Ready.";

		}

		private void addDocFileButton_Click(object sender, RoutedEventArgs e)
		{
			
			System.Windows.Forms.OpenFileDialog fileBrowser = new System.Windows.Forms.OpenFileDialog();
			fileBrowser.Filter = "Documents (*.doc;*.docx)|*.doc;*.docx|All files (*.*)|*.*" ;

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
					statusLabel.Content = "Error: Could not read file. Original error: " + ex.Message;
					return;
				}
				// check if have merge doc
				if (!checkWord())
				{
					addDocLocationLabel.Content = "";
					return;
				}
				statusLabel.Content = "Click Yes in word...";
				//ghetto to make the label update
				Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background,
										  new Action(delegate { }));

				Word.Document doc = wordApp.Application.Documents.Open(fileBrowser.FileName, ReadOnly:true, Visible:false);
				//Console.WriteLine(doc.MailMerge.State);
				//Console.WriteLine(doc.MailMerge.DataSource.ConnectString);
				//Console.WriteLine(doc.MailMerge.DataSource.QueryString);
				if (doc.MailMerge.State == Word.WdMailMergeState.wdMainAndDataSource)
				{
					addDocLocationLabel.Content = fileBrowser.FileName;
					statusLabel.Content = "Ready.";
					//Console.WriteLine("" + doc.MailMerge.DataSource.DataFields[1].Value + doc.MailMerge.DataSource.DataFields[2].Value +
					//doc.MailMerge.DataSource.DataFields[3].Value);
				}
				else
				{
					addDocLocationLabel.Content = "";
					statusLabel.Content = "Selected document does not have merge data.";
				}
				((Word._Document)doc).Close(SaveChanges:false);
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
			mergedocs.Clear();
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
				statusLabel.Content = e.UserState as String;
			progress.Value = e.ProgressPercentage;
			
		}

		private void processMergeWorker(object sender, DoWorkEventArgs e)
		{
			BackgroundWorker worker = sender as BackgroundWorker;
			mergedocs.Clear();
			try { Directory.Delete(tempMerge, true); }
			catch (System.IO.DirectoryNotFoundException) { }
			Directory.CreateDirectory(tempMerge);

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
					for (int x = 0; x < 100; x++ )
					{
						try
						{
							string col = headrow[x].ToString().Trim().ToLower();
							if (col.Equals("")) { continue; }
							headers.Add(col, x);
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
						for (int xi = 1; xi <= HASHFIELDNUMS; xi++)
						{
							String oledata = processFloat(olerow[xi - 1].ToString());
							try { sb.Append(oledata); }
							catch (COMException) { continue; }
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
						mergedocs[hash].Add(new string[] { docname, attachname });

						if (File.Exists(docname))
						{
							// bail and show error
							Logger.log(TraceEventType.Verbose, 99, "Non unique: " + docname + "\r\n" + "hash:" + hash + " hashed data: " + sb.ToString()); 
							doc.Close(SaveChanges: false);
							mergedocs.Clear();
							throw new DataException("IMPORTANT ERROR: Data source has non unique data.");
						}
						d.attachmentName = docname;
						doc.MailMerge.Execute(Pause: false);
						Word._Document nd = wordApp.ActiveDocument;
						nd.ExportAsFixedFormat(OutputFileName: docname,
							ExportFormat: Word.WdExportFormat.wdExportFormatPDF);
						nd.Close(SaveChanges: false);

						int skips = 0;
						do
						{
							try
							{
								doc.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdNextRecord;
							}
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
				Directory.Delete(tempDirectory, true);
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
					statusLabel.Content = "Invalid column name.";
					return null;
				}
				matched = true;
				if (headers.ContainsKey(field)) {
					newaname = newaname.Replace(match.Value, row[headers[field]].ToString());
				} else {
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
			else
			{
				statusLabel.Content = "Merges not created yet.";
			}
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

	}
}
