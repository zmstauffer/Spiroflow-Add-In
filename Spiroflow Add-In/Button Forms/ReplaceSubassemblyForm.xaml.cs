using Inventor;
using SpiroflowVault;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SpiroflowAddIn.Button_Forms
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
			WindowStartupLocation = WindowStartupLocation.CenterScreen;
		}

		public void OnItemMouseDoubleClick(object sender, MouseButtonEventArgs args)
		{
			if (sender is TreeViewItem && !((TreeViewItem)sender).IsSelected) return;

			var item = (TreeViewItem)sender;
			{
				if (item.DataContext.GetType() == typeof(FolderInfo)) return;
			}

			var file = (VaultFileInfo)item.Header;

			if (!System.IO.File.Exists(file.LocalFilePath))
			{
				VaultFunctions.DownloadFileById(file.Id);
			}

			subAssyToReplace.Replace(file.LocalFilePath, false);

			this.Close();
		}

		private void CloseCommandBinding_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs eventArgs)
		{
			this.Close();
		}
	}
}
