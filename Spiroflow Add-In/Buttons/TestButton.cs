using Inventor;
using SpiroflowAddIn.Utilities;
using System;
using System.Collections.Generic;
using System.Windows;
using SpiroflowAddIn.Button_Forms;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Buttons
{
	public class TestButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public TestButton()
		{
			DisplayName = "Test Button";
			InternalName = "testButton";
			PanelID = "assemblyModelPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			var testForm = new TestForm();
			testForm.ShowDialog();
		}

		//private void SetOccurrencePartNumbers(ComponentOccurrences occurrences)
		//{
		//	foreach (ComponentOccurrence occurrence in occurrences)
		//	{
		//		var nodeName = occurrence.Name.Split(":".ToCharArray(), StringSplitOptions.None);
		//		var instanceNum = nodeName.Length > 1 ? $":{nodeName[1]}" : ":1";
		//		try
		//		{
		//			occurrence.Name = $"{GetPartNumber(occurrence)}{instanceNum}";
		//		}
		//		catch (Exception ex)
		//		{
		//			MessageBox.Show($"Test Button Error: {ex}");
		//			continue;
		//		}
		//		if (occurrence.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject) SetOccurrencePartNumbers(occurrence.Definition.Occurrences);
		//	}
		//}

		//private string GetPartNumber(ComponentOccurrence occurrence)
		//{
		//	var propertySet = occurrence.Definition.Document.PropertySets["Design Tracking Properties"];
		//	string partNumber = (string)propertySet["Part Number"].Value;
		//	return partNumber;
		//}
	}
}
