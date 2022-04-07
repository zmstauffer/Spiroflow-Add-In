using Inventor;
using SpiroflowAddIn.Utilities;
using System.Windows;

namespace SpiroflowAddIn.Buttons
{
	public class ChangePartNumberToFilenameButton : IButton
	{
		public Inventor.Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }
		
		public ChangePartNumberToFilenameButton()
		{
			DisplayName = $"Change Part #{System.Environment.NewLine}To Filename";
			InternalName = "changePartNumToFilename";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.changePartNumber);
		}

		public void Execute(NameValueMap context)
		{
			Document doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			int filesAvailableForEditing = 0;
			int currentFileCount = 0;
			
			//capture how many files are read-only because we are not going to modify those.
			foreach (Document referencedDoc in doc.AllReferencedDocuments)
			{
				System.IO.FileInfo fileInfo = new System.IO.FileInfo(referencedDoc.FullFileName);
				if (!fileInfo.IsReadOnly) filesAvailableForEditing++;
			}

			if (filesAvailableForEditing <= 0)			//nothing to do here
			{
				MessageBox.Show("All files in this assembly are read only.");
				return;
			}

			var progressBar = invApp.CreateProgressBar(false, filesAvailableForEditing, "Changing Part Numbers...", true);
			progressBar.Message = "Changing Part Numbers...";

			foreach (Document file in doc.AllReferencedDocuments)
			{
				System.IO.FileInfo fileInfo = new System.IO.FileInfo(file.FullFileName);
				if (!fileInfo.IsReadOnly)
				{
					currentFileCount++;
					progressBar.Message = $"Changing Part Number for file {currentFileCount} of {filesAvailableForEditing}";
					progressBar.UpdateProgress();

					invApp.Documents.Open(file.FullFileName, false);

					string docName = System.IO.Path.GetFileNameWithoutExtension(file.FullFileName);

					file.DisplayName = docName;
					file.PropertySets["Design Tracking Properties"]["Part Number"].Value = docName;

					invApp.SilentOperation = true;
					file.Close(false);
					invApp.SilentOperation = false;
				}
			}
			progressBar.Close();
		}
	}
}
