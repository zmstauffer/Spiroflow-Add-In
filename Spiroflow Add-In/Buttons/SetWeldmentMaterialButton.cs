using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Interop;
using Inventor;
using SpiroflowAddIn.Button_Forms;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;
using Application = Inventor.Application;
using IPictureDisp = stdole.IPictureDisp;

namespace SpiroflowAddIn.Buttons
{
	public class SetWeldmentMaterialButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		private string subAssemblyPath { get; set; }
		private List<FolderInfo> folders { get; set; }

		public SetWeldmentMaterialButton()
		{
			DisplayName = $"Set Weld{System.Environment.NewLine}Material";
			InternalName = "setWeldMaterial";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.weldMaterial, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.weldMaterial, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			try
			{
				AssemblyDocument assyDoc = (AssemblyDocument) doc;
				var weldmentDef = (WeldmentComponentDefinition) assyDoc.ComponentDefinition;

				if (weldmentDef != null)
				{
					if (assyDoc.FullFileName.Contains("-CS"))
					{
						weldmentDef.WeldBeadMaterial = assyDoc.MaterialAssets["STEEL, CARBON"];
					}
					else if (assyDoc.FullFileName.Contains("-304"))
					{
						weldmentDef.WeldBeadMaterial = assyDoc.MaterialAssets["STAINLESS STEEL 304"];
					}
					else if (assyDoc.FullFileName.Contains("-316"))
					{
						weldmentDef.WeldBeadMaterial = assyDoc.MaterialAssets["STAINLESS STEEL 316"];
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				return;
			}

		}
	}
}