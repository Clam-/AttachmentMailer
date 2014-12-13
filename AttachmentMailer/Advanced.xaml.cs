using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AttachmentMailer
{
	/// <summary>
	/// Interaction logic for Advanced.xaml
	/// </summary>
	public partial class Advanced : Window
	{
		public Advanced()
		{
			InitializeComponent();
			allownonunique.IsChecked = Option.allowduplicatehash;
			uniqueOnly.IsChecked = Option.createforuniquehash;
			hashColumnsText.Text = Option.hashcolumns;
		}

		private void saveButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				Option.getColumns(hashColumnsText.Text);
			}
			catch (Exception)
			{
				statusLabel.Content = "Invalid column list. (1,3,4)";
				return;
			}
			Option.allowduplicatehash = allownonunique.IsChecked.Value;
			Option.createforuniquehash = uniqueOnly.IsChecked.Value;
			Option.hashcolumns = hashColumnsText.Text;
			Close();
		}

		private void cancelButton_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}
	}
}
