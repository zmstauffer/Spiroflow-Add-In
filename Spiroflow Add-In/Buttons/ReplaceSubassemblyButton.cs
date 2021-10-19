using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using SpiroflowAddIn.Utilities;
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

		public ReplaceSubassemblyButton()
		{
			DisplayName = "Replace Subassembly";
			InternalName = "replaceSubassembly";
			PanelID = "bomPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			throw new NotImplementedException();
		}
	}
}
