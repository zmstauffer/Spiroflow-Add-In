using Autodesk.iLogic.Automation;
using Inventor;
using SpiroflowAddIn.Button_Forms;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using Application = Inventor.Application;
using Path = System.IO.Path;

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
		private List<string> fileList { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }

		public FindMissingFabricationDrawingsButton()
		{
			DisplayName = $"Find Missing{System.Environment.NewLine}Fab Drawings";
			InternalName = "findMissingDrawings";
			PanelID = "assemblyBookkeepingPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.findMissingFabDrawings, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.findMissingFabDrawings, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;
			var neededFiles = new List<string>();
			fileList = new List<string>();

			if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assemblyDoc = (AssemblyDocument)doc;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				var filename = occurrence.ReferencedDocumentDescriptor.DisplayName;
				filename = Path.ChangeExtension(filename, ".idw");
				fileList.Add(filename);
				if (occurrence.BOMStructure != BOMStructureEnum.kInseparableBOMStructure) GetSubOccurrences(occurrence);
			}

			var existingFileList = GetExistingFiles();

			if (existingFileList == null || existingFileList.Count < 1) return;

			foreach (var name in fileList)
			{
				//see if file has drawing in vault
				var vaultFileList = VaultFunctions.FindFilesByFilename(name);

				if (vaultFileList.Count > 0)
				{
					var file = existingFileList.FirstOrDefault(x => x.Contains(name));
					if (file is null) neededFiles.Add(name);
				}
			}

			neededFiles = neededFiles.Distinct().ToList();

			var sticker = neededFiles.FirstOrDefault(x => x.Contains("STICKER"));
			if (sticker != null) neededFiles.Remove(sticker);

			if (neededFiles.Count <= 0)
			{
				MessageBox.Show("No files need to be printed.");
				return;
			}

			var form = new FindMissingFabDrawingsForm();
			var helper = new WindowInteropHelper(form);
			helper.Owner = new IntPtr(invApp.MainFrameHWND);

			form.fileList.ItemsSource = neededFiles;
			var dialogResult = form.ShowDialog();

			if (dialogResult.HasValue && dialogResult.Value)
			{
				ExportToDwg(neededFiles);
			}
		}

		private void ExportToDwg(List<string> neededFiles)
		{
			List<string> filenames = new List<string>();

			foreach (var file in neededFiles)
			{
				filenames.Add(file + ".idw");
			}

			var progressBar = invApp.CreateProgressBar(false, filenames.Count, "Creating DWGs...");
			var currentStep = 1;

			foreach (var filename in filenames)
			{
				progressBar.Message = $"Creating DWG {currentStep} of {filenames.Count} - {filename}";
				progressBar.UpdateProgress();

				var files = VaultFunctions.FindFilesByFilename(filename);
				if (files.Count != 1)
				{
					MessageBox.Show($"File {filename} was found multiple times in Vault, or not found at all. Please print that one yourself.");
				}
				else
				{
					//download file
					VaultFunctions.DownloadFileById(files[0].Id);

					var ilogicAddin = invApp.ApplicationAddIns.ItemById["{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}"];
					iLogicAutomation ilogicAutomation = (iLogicAutomation)ilogicAddin.Automation;

					//open file, but disable ilogic first. This is because some drawings have an "Update" rule that requires the model to update but fails.
					try
					{
						ilogicAutomation.RulesOnEventsEnabled = false;
						var doc = invApp.Documents.Open(files[0].LocalFilePath, false);

						DrawingDocument drawingDoc = (DrawingDocument)doc;

						DWGPrinter.Print(drawingDoc, invApp);
						doc.Close(true);
					}
					catch { }
					finally
					{
						ilogicAutomation.RulesOnEventsEnabled = true;
					}
				}
				progressBar.UpdateProgress();
				currentStep++;
			}

			progressBar.Close();
		}

		private void GetSubOccurrences(ComponentOccurrence occurrence)
		{
			foreach (ComponentOccurrence subOccurrence in occurrence.Definition.Occurrences)
			{
				if (subOccurrence.Name.Contains("SP") || subOccurrence.Name.Contains("PC") || subOccurrence.Name.Contains("CF") && occurrence.Name.Count(x => x == '-') <= 1)
				{
					if (!fileList.Contains(subOccurrence.Name)) fileList.Add(subOccurrence.Name);
				}
				if (subOccurrence.BOMStructure != BOMStructureEnum.kInseparableBOMStructure) GetSubOccurrences(subOccurrence);
			}
		}

		private List<string> GetExistingFiles()
		{
			var manufacturingDrawingFilePath = SettingService.GetSetting("ManufacturingDrawingsPath");
			try
			{
				var files = Directory.GetFiles(manufacturingDrawingFilePath, "*.dwg", SearchOption.AllDirectories).ToList();
				return files;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error reading manufacturing files. Please check directory. Error: {ex.Message}");
				return null;
			}
		}
	}
}
