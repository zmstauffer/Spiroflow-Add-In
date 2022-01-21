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

namespace SpiroflowViewModel.Button_Forms
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
		}

		private void Window_Loaded(object sender, RoutedEventArgs eventArgs)
		{
			Top = Mouse.GetPosition(null).Y;
			Left = Mouse.GetPosition(null).X;
		}

		private void CloseCommandBinding_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs eventArgs)
		{
			if(eventArgs.OriginalSource.Equals(Yes)) DialogResult = true;
			if (eventArgs.OriginalSource.Equals(No)) DialogResult = false;
			this.Close();
		}
	}
}
