using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentMailer
{
	public static class Option
	{
		public static Boolean allowduplicatehash = false;
		public static Boolean createforuniquehash = false;
		public static String hashcolumns = "";

		public static int[] getColumns()
		{
			return getColumns(hashcolumns);
		}
		public static int[] getColumns(String s)
		{
			if (s.Equals("")) { return null; }
			string[] sa = s.Split(',');
			int[] ia = new int[sa.Length];
			for (int x = 0; x < sa.Length; x++)
			{
				ia[x] = Convert.ToInt32(sa[x]);
			}
			return ia;
		}
	}
	
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
		//public String attachmentName { get; set; }

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
