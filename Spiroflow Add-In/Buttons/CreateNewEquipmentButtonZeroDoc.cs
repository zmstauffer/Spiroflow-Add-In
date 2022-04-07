using Inventor;
using SpiroflowAddIn.Utilities;
//using SpiroflowAddIn.WPF_Forms;

namespace SpiroflowAddIn.Buttons
{
	class CreateNewEquipmentButtonZeroDoc : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public CreateNewEquipmentButtonZeroDoc()
		{
			DisplayName = $"Create New{System.Environment.NewLine}Equipment";
			InternalName = "createNewEquipmentZeroDoc";
			PanelID = "spiroflowPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			//CreateNewEquipmentForm createNewEquipmentForm = new CreateNewEquipmentForm();
			//createNewEquipmentForm.ShowDialog();
		}
	}


}
