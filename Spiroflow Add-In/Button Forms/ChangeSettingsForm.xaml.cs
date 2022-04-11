using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Button = System.Windows.Controls.Button;
using SpiroflowAddIn.Utilities;
using System.Windows.Forms;
using System.Windows.Navigation;

namespace SpiroflowAddIn.Button_Forms
{
	/// <summary>
	/// Interaction logic for UserControl1.xaml
	/// </summary>
	public partial class ChangeSettingsForm : Window, INotifyPropertyChanged
	{

		public event PropertyChangedEventHandler PropertyChanged;
		
		private string _BOMExportPath;
		public string BOMExportPath
		{
			get => _BOMExportPath;
			set
			{
				if (_BOMExportPath != value)
				{
					_BOMExportPath = value;
					RaisePropertyChanged(nameof(BOMExportPath));
				}
			}
		}

		private string _DrawingExportPath { get; set; }
		public string DrawingExportPath
		{
			get => _DrawingExportPath;
			set
			{
				if (_DrawingExportPath != value)
				{
					_DrawingExportPath = value;
					RaisePropertyChanged(nameof(DrawingExportPath));
				}
			}
		}

		public ChangeSettingsForm()
		{
			InitializeSettingFields();
			InitializeComponent();
			WindowStartupLocation = WindowStartupLocation.CenterScreen;
		}

		private void BOMExportButton_OnClick(object sender, RoutedEventArgs e)
		{
			var dialog = new FolderBrowserDialog();
			var result = dialog.ShowDialog();

			if (result == System.Windows.Forms.DialogResult.OK) BOMExportPath = dialog.SelectedPath;
		}

		private void DrawingExportButton_OnClick(object sender, RoutedEventArgs e)
		{
			var dialog = new FolderBrowserDialog();
			var result = dialog.ShowDialog();

			if (result == System.Windows.Forms.DialogResult.OK) DrawingExportPath = dialog.SelectedPath;
		}

		private void SaveButton_OnClick(object sender, RoutedEventArgs e)
		{
			if (SettingService.GetSetting("BOMExportPath") != BOMExportPath) SettingService.SetSetting("BOMExportPath", BOMExportPath);
			if (SettingService.GetSetting("DrawingExportPath") != DrawingExportPath) SettingService.SetSetting("DrawingExportPath", DrawingExportPath);

			CloseCommandBinding_Executed(sender, null);
		}

		private void CancelButton_OnClick(object sender, RoutedEventArgs e)
		{
			CloseCommandBinding_Executed(sender, null);
		}

		private void CloseCommandBinding_Executed(object sender, ExecutedRoutedEventArgs eventArgs)
		{
			this.Close();
		}

		private void InitializeSettingFields()
		{
			BOMExportPath = SettingService.GetSetting("BOMExportPath");
			DrawingExportPath = SettingService.GetSetting("DrawingExportPath");
		}

		private void RaisePropertyChanged(string property)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
		}
	}
}
