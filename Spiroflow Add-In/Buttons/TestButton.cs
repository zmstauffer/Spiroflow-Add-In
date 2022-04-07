using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using Inventor;
using SpiroflowAddIn.Utilities;
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
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			if (invApp.ActiveDocument.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assemblyDoc = (AssemblyDocument) invApp.ActiveDocument;
			var assemblyDef = assemblyDoc.ComponentDefinition;

			foreach (ComponentOccurrence occurrence in assemblyDef.Occurrences)
			{
				var nodeName = occurrence.Name.Split(":".ToCharArray(), StringSplitOptions.None);
				var instanceNum = nodeName.Length > 1? $":{nodeName[1]}" : ":1";
				try
				{
					occurrence.Name = $"{GetPartNumber(occurrence)}{instanceNum}";
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Test Button Error: {ex}");
					continue;
				}
			}
		}

		private string GetPartNumber(ComponentOccurrence occurrence)
		{
			var propertySet = occurrence.Definition.Document.PropertySets["Design Tracking Properties"];
			string partNumber = (string)propertySet["Part Number"].Value;
			return partNumber;
		}
	}
}
