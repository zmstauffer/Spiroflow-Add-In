using Inventor;
using SpiroflowAddIn.Utilities;
//using SpiroflowAddIn.WPF_Forms;
using System;

namespace SpiroflowAddIn.Buttons
{
	public class CreateNewEquipmentButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public CreateNewEquipmentButton()
		{
			DisplayName = $"Create New{System.Environment.NewLine}Equipment";
			InternalName = "createNewEquipment";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			//CreateNewEquipmentForm createNewEquipmentForm = new CreateNewEquipmentForm();
			//createNewEquipmentForm.ShowDialog();
		}
	}
}
