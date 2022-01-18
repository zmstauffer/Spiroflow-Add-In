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

		private void CloseCommandBinding_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs eventArgs)
		{
			result = false;
			this.Close();
		}
	}
}
