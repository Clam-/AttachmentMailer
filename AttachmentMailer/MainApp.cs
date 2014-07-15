using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentMailer
{
	public partial class MainApp : System.Windows.Application
	{
		private static TraceSource logging =
			new TraceSource("AttachmentMailer");
		
		/// <summary>
		/// InitializeComponent
		/// </summary>
		public void InitializeComponent()
		{

		#line 4 "App.xaml"
			this.StartupUri = new System.Uri("MainWindow.xaml", System.UriKind.Relative);

		#line default
		#line hidden
		}

		/// <summary>
		/// Application Entry Point.
		/// </summary>
		[System.STAThreadAttribute()]
		[System.Diagnostics.DebuggerNonUserCodeAttribute()]
		[System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
		public static void Main()
		{
			// http://stackoverflow.com/a/10203030
			AppDomain currentDomain = default(AppDomain);
			currentDomain = AppDomain.CurrentDomain;
			// Handler for unhandled exceptions.
			currentDomain.UnhandledException += GlobalUnhandledExceptionHandler;
			// Handler for exceptions in threads behind forms.
			System.Windows.Forms.Application.ThreadException += GlobalThreadExceptionHandler;

			AttachmentMailer.App app = new AttachmentMailer.App();
			app.InitializeComponent();
			app.Run();
		}

		// http://stackoverflow.com/a/10203030
		private static void GlobalUnhandledExceptionHandler(object sender, UnhandledExceptionEventArgs e)
		{
			Exception ex = default(Exception);
			ex = (Exception)e.ExceptionObject;
			logging.TraceEvent(TraceEventType.Critical, 9, "\r\n" + DateTime.Now.ToString() +
				"\r\nCRASH\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace);
		}

		private static void GlobalThreadExceptionHandler(object sender, System.Threading.ThreadExceptionEventArgs e)
		{
			Exception ex = default(Exception);
			ex = e.Exception;
			logging.TraceEvent(TraceEventType.Critical, 10, "\r\n" + DateTime.Now.ToString() + 
				"\r\nTHREAD CRASH\r\n" + ex.GetType() + ":" + ex.Message + "\r\n" + ex.StackTrace + "\r\n");
		}
	}
}
