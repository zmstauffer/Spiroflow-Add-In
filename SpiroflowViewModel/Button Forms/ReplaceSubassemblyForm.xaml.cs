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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Inventor;
using Spiroflow_Vault;
using SpiroflowVault;

namespace SpiroflowViewModel.Button_Forms
{
	/// <summary>
	/// Interaction logic for ReplaceSubassemblyForm.xaml
	/// </summary>
	public partial class ReplaceSubassemblyForm : Window
	{

		public ComponentOccurrence subAssyToReplace;

		public ReplaceSubassemblyForm()
		{
			InitializeComponent();
		}

		public void OnItemMouseDoubleClick(object sender, MouseButtonEventArgs args)
		{
			if (sender is TreeViewItem && !((TreeViewItem)sender).IsSelected) return;

			var item = (TreeViewItem) sender;
			{
				if (item.DataContext.GetType() == typeof(FolderInfo)) return;
			}

			VaultFileInfo file = (VaultFileInfo)item.Header;

			if (!System.IO.File.Exists(file.LocalFilePath))
			{
				VaultFunctions.DownloadFileById(file.Id);
			}

			subAssyToReplace.Replace(file.LocalFilePath, false);

			this.Close();

		}
	}
}
