using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	public interface IButton
	{
		Application invApp { get; set; }
		string DisplayName { get; set; }
		string InternalName { get; set; }
		string PanelID { get; set; }
		stdole.IPictureDisp icon { get; set; }
		ButtonDefinition buttonDef { get; set; }

		void Execute(NameValueMap context);
	}
}
