using System.Windows;
using System.Windows.Input;

namespace SpiroflowAddIn.Button_Forms
{
	/// <summary>
	/// Interaction logic for FindMissingFabDrawingsForm.xaml
	/// </summary>
	public partial class FindMissingFabDrawingsForm : Window
	{
		public bool? result { get; set; }
		public FindMissingFabDrawingsForm()
		{
			InitializeComponent();
			WindowStartupLocation = WindowStartupLocation.CenterScreen;
		}

		private void CloseCommandBinding_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs eventArgs)
		{
			if(eventArgs.OriginalSource.Equals(Yes)) DialogResult = true;
			if (eventArgs.OriginalSource.Equals(No)) DialogResult = false;
			this.Close();
		}
	}
}
