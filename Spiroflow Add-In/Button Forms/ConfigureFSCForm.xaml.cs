using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Inventor;
using SpiroflowAddIn.Buttons;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Button_Forms
{
	/// <summary>
	/// Interaction logic for ConfigureFSCForm.xaml
	/// </summary>
	public partial class ConfigureFSCForm : Window
	{
		public ConfigureFSCForm(ConfigureFSCButton _dataContext)
		{
			this.DataContext = _dataContext;
			InitializeComponent();
		}

		private void CloseCommandBinding_Executed(object sender, ExecutedRoutedEventArgs eventArgs)
		{
			this.Close();
			GetInventorApp.GetApp().ActiveDocument.Update2();
		}

		private void TypeComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			var button = (ConfigureFSCButton)DataContext;
			button.TypeChanged();
		}

		private void LengthTextBox_OnKeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Return)
			{
				var assemblyDocument = (AssemblyDocument)GetInventorApp.GetApp().ActiveDocument;
				var unitsOfMeasure = assemblyDocument.UnitsOfMeasure;

				string newLength = "";
				try
				{
					newLength = $"{(unitsOfMeasure.GetValueFromExpression(LengthTextBox.Text, Inventor.UnitsTypeEnum.kInchLengthUnits) / 2.54).ToString()} in";
				}
				catch
				{
					//if expression doesn't make sense, just get length of assembly again
					var button = (ConfigureFSCButton)DataContext;
					var lengthInCentimeters = button.assemblyDoc.ComponentDefinition.Parameters["length"].Value;
					newLength = button.ConvertLengthToInches((lengthInCentimeters / 2.54).ToString());
				}

				LengthTextBox.Text = newLength; 
			}
		}
	}
}
