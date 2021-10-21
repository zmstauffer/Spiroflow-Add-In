using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Interop;
using Inventor;
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
		private TreeView fileTreeView { get; set; }

		public ReplaceSubassemblyButton()
		{
			DisplayName = "Replace Subassembly";
			InternalName = "replaceSubassembly";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			ComponentOccurrence subAssemblyToReplace = (ComponentOccurrence)invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select Subassembly to Replace.");

			if (subAssemblyToReplace is null) return;

			Document subAssyDoc = (Document)subAssemblyToReplace.Definition.Document;
			var subAssemblyPath = subAssyDoc.FullFileName;

			//open form
			var form = new ReplaceSubassemblyForm();
			var helper = new WindowInteropHelper(form);
			helper.Owner = new IntPtr(invApp.MainFrameHWND);

			fileTreeView = new TreeView();

			//form.Con
			form.ShowDialog();

			invApp.ActiveDocument.Update();
		}

		private void InitializeTreeView()
		{
			List<FolderInfo> subFolders = VaultFunctions.GetFolderNames(subAssemblyPath);

			if (subFolders == null) return;



		}
	}
}
