using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Interop;
using Inventor;
using Spiroflow_Vault;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;
using SpiroflowViewModel.Button_Forms;
using IPictureDisp = stdole.IPictureDisp;

namespace SpiroflowAddIn.Buttons
{
	class ReplaceSubassemblyButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public IPictureDisp icon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		private string subAssemblyPath { get; set; }
		private List<FolderInfo> folders { get; set; }


		public ReplaceSubassemblyButton()
		{
			DisplayName = $"Replace{System.Environment.NewLine}Subassembly";
			InternalName = "replaceSubassembly";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			ComponentOccurrence subAssemblyToReplace;
			if (invApp.ActiveDocument.SelectSet.Count <= 0 || invApp.ActiveDocument.SelectSet.Count > 1)				//check if we already have something selected, make sure it's only 1 thing
			{
				subAssemblyToReplace = (ComponentOccurrence) invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select Subassembly to Replace.");
			}
			else
			{
				subAssemblyToReplace = (ComponentOccurrence)invApp.ActiveDocument.SelectSet[1];				//index starts at 1...
			}

			if (subAssemblyToReplace is null) return;

			Document subAssyDoc = (Document)subAssemblyToReplace.Definition.Document;
			
			subAssemblyPath = subAssyDoc.FullFileName;
			subAssemblyPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(subAssemblyPath, @"..\..\"));
			
			//open form
			var form = new ReplaceSubassemblyForm();
			var helper = new WindowInteropHelper(form);
			helper.Owner = new IntPtr(invApp.MainFrameHWND);

			InitializeTreeView();

			form.fileTreeView.ItemsSource = folders;
			form.subAssyToReplace = subAssemblyToReplace;

			form.ShowDialog();

			invApp.ActiveDocument.Update();
		}

		private void InitializeTreeView()
		{
			folders = VaultFunctions.GetFolderNames(subAssemblyPath);

			if (folders == null || folders.Count == 0) return;

			foreach (var folder in folders)
			{
				folder.files = VaultFunctions.GetFilenamesFromFolderId(folder.folderID);
			}

			folders.RemoveAll(x => x.files.Count == 0);	//remove folders w/ no files, as you can't pick that anyways
		}
	}
}
