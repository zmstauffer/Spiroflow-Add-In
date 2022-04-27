using SpiroflowAddIn.Utilities;
using System.ComponentModel;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

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

		private string _ManufacturingDrawingsPath { get; set; }
		public string ManufacturingDrawingsPath
		{
			get => _ManufacturingDrawingsPath;
			set
			{
				if (_ManufacturingDrawingsPath != value)
				{
					_ManufacturingDrawingsPath = value;
					RaisePropertyChanged(nameof(ManufacturingDrawingsPath));
				}
			}
		}

		private string _Engineer { get; set; }
		public string Engineer
		{
			get => _Engineer;
			set
			{
				if (_Engineer != value)
				{
					_Engineer = value;
					RaisePropertyChanged(nameof(Engineer));
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

			if (result == System.Windows.Forms.DialogResult.OK) BOMExportPath = $@"{dialog.SelectedPath}";
			if (BOMExportPath.Substring(BOMExportPath.Length - 1) != @"\") BOMExportPath += @"\";               //have to make sure file has trailing slash, as all future uses of this will have it
		}

		private void DrawingExportButton_OnClick(object sender, RoutedEventArgs e)
		{
			var dialog = new FolderBrowserDialog();
			var result = dialog.ShowDialog();

			if (result == System.Windows.Forms.DialogResult.OK) DrawingExportPath = $@"{dialog.SelectedPath}";
			if (DrawingExportPath.Substring(DrawingExportPath.Length - 1) != @"\") DrawingExportPath += @"\";
		}

		private void ManufacturingDrawingButton_OnClick(object sender, RoutedEventArgs e)
		{
			var dialog = new FolderBrowserDialog();
			var result = dialog.ShowDialog();

			if (result == System.Windows.Forms.DialogResult.OK) ManufacturingDrawingsPath = $@"{dialog.SelectedPath}";
			if (ManufacturingDrawingsPath.Substring(ManufacturingDrawingsPath.Length - 1) != @"\") ManufacturingDrawingsPath += @"\";
		}

		private void SaveButton_OnClick(object sender, RoutedEventArgs e)
		{
			if (SettingService.GetSetting("BOMExportPath") != BOMExportPath) SettingService.SetSetting("BOMExportPath", BOMExportPath);
			if (SettingService.GetSetting("DrawingExportPath") != DrawingExportPath) SettingService.SetSetting("DrawingExportPath", DrawingExportPath);
			if (SettingService.GetSetting("ManufacturingDrawingsPath") != DrawingExportPath) SettingService.SetSetting("ManufacturingDrawingsPath", ManufacturingDrawingsPath);
			if (SettingService.GetSetting("Engineer") != Engineer) SettingService.SetSetting("Engineer", Engineer);

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
			ManufacturingDrawingsPath = SettingService.GetSetting("ManufacturingDrawingsPath");
			Engineer = SettingService.GetSetting("Engineer");
		}

		private void RaisePropertyChanged(string property)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
		}

	}
}
