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
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Button_Forms
{
	/// <summary>
	/// Interaction logic for UserControl1.xaml
	/// </summary>
	public partial class TestForm : Window
	{

		public string TestText { get; set; }

		public TestForm()
		{
			InitializeComponent();
		}

		private void TestTextBox_OnKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Return)
			{
				var assemblyDocument = (AssemblyDocument)GetInventorApp.GetApp().ActiveDocument;
				var unitsOfMeasure = assemblyDocument.UnitsOfMeasure;

				TestTextBox.Text = (unitsOfMeasure.GetValueFromExpression(TestTextBox.Text, Inventor.UnitsTypeEnum.kInchLengthUnits) / 2.54).ToString() + "in";
			}
		}

		private void CloseCommandBinding_Executed(object sender, ExecutedRoutedEventArgs eventArgs)
		{
			this.Close();
		}
	}
}
