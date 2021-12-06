using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Inventor;
using SpiroflowAddIn.Utilities;
using Application = Inventor.Application;
using Environment = Inventor.Environment;

namespace SpiroflowAddIn.Buttons
{
	class FindMissingFabricationDrawingsButton : IButton
	{
		public Application invApp { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public ButtonDefinition buttonDef { get; set; }
		private List<string> fileList { get; }

		public FindMissingFabricationDrawingsButton()
		{
			DisplayName = $"Find Missing{System.Environment.NewLine}Fab Drawings";
			InternalName = "findMissingDrawings";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
			fileList = new List<string>();
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;
			var neededFiles = new List<string>();
			
			if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assemblyDoc = (AssemblyDocument)doc;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				if (occurrence.Name.Contains("SP") || occurrence.Name.Contains("PC") || occurrence.Name.Contains("CF"))
				{
					fileList.Add(occurrence.Name);
					
				}
				if (occurrence.BOMStructure != BOMStructureEnum.kInseparableBOMStructure) GetSubOccurrences(occurrence);
			}

			var existingFileList = GetExistingFiles();

			foreach (var name in fileList)
			{
				var filename = name.Split(new[] { ':' })[0];
				var file = existingFileList.FirstOrDefault(x => x.Contains(filename));
				if (file is null) neededFiles.Add(filename);
			}

			neededFiles = neededFiles.Distinct().ToList();

			var outputString = string.Join(System.Environment.NewLine, neededFiles);
			if (outputString == "") outputString = "No missing drawings.";
			MessageBox.Show(outputString);
		}

		private void GetSubOccurrences(ComponentOccurrence occurrence)
		{
			foreach (ComponentOccurrence subOccurrence in occurrence.Definition.Occurrences)
			{
				if (subOccurrence.Name.Contains("SP") || subOccurrence.Name.Contains("CF") && (subOccurrence.Name.Contains("-CS") || subOccurrence.Name.Contains("-304")))
				{
					if(!fileList.Contains(subOccurrence.Name)) fileList.Add(subOccurrence.Name);
				}
				if (subOccurrence.BOMStructure != BOMStructureEnum.kInseparableBOMStructure) GetSubOccurrences(subOccurrence);
			}
		}

		private List<string> GetExistingFiles()
		{
			var manufacturingDrawingFilePath = @"C:\Users\zstauffer\Spiroflow Systems\Spiroflow Systems Team Site - Job Files - 2018 onward\Manufacturing Drawings";
			return Directory.GetFiles(manufacturingDrawingFilePath, "*.dwg", SearchOption.AllDirectories).ToList();
		}
	}
}
