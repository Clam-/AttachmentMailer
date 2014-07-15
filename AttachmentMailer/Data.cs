using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentMailer
{
	
	public class Attachment
	{
		public String location { get; set; }
		public String attachmentName { get; set; }

		public Attachment() { }
		public Attachment(String location, String attachmentName)
		{
			this.location = location;
			this.attachmentName = attachmentName;
		}
	}

	public class Document
	{
		public String location { get; set; }
		public String attachmentFormat { get; set; }
		public String attachmentName { get; set; }

		public Document() { }
		public Document(String location, String attachmentFormat)
		{
			this.location = location;
			this.attachmentFormat = attachmentFormat;
		}
	}

	public class Replacement
	{
		public string placeholder { get; set; }
		public int replacement { get; set; }

		public Replacement() { }
		public Replacement(string placeholder, int replacement)
		{
			this.placeholder = placeholder;
			this.replacement = replacement;
		}
	}

}
