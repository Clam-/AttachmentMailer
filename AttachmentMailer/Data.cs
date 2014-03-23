using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttachmentMailer
{
	
	public class Data
	{
		public String location { get; set; }
		public String attachmentName { get; set; }

		public Data() { }
		public Data(String location, String attachmentName)
		{
			this.location = location;
			this.attachmentName = attachmentName;
		}
	}

	public class Datum : ObservableCollection<Data>
	{
		public Datum() { }
	}

	public class DataReplace
	{
		public string placeholder { get; set; }
		public int replacement { get; set; }

		public DataReplace() { }
		public DataReplace(string placeholder, int replacement)
		{
			this.placeholder = placeholder;
			this.replacement = replacement;
		}
	}

	public class DataReplacements : ObservableCollection<DataReplace>
	{
		public DataReplacements() { }
	}
}
