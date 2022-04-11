using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Inventor;
using SpiroflowAddIn.Button_Forms;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;

namespace SpiroflowAddIn.Buttons
{
	public class ShowSettingsWindowButton : IButton
	{
		public Inventor.Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public ShowSettingsWindowButton()
		{
			DisplayName = $"Show Settings";
			InternalName = "showSettings";
			PanelID = "spiroflowPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.settings, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.settings, 16, 16));
		}
		public void Execute(NameValueMap context)
		{
			//open form
			var form = new ChangeSettingsForm();
			var helper = new WindowInteropHelper(form);
			helper.Owner = new IntPtr(invApp.MainFrameHWND);

			form.ShowDialog();
		}

	}
}

