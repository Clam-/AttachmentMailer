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

		public Data()
		{

		}
		public Data(String location, String attachmentName)
		{
			this.location = location;
			this.attachmentName = attachmentName;
		}

	}

	public class Datum : ObservableCollection<Data>
	{
		public Datum()
		{
		}

	}
}
