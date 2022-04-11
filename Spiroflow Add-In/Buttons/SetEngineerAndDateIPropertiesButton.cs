using Inventor;
using SpiroflowAddIn.Utilities;
using System;
using System.Windows;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Buttons
{
	public class SetEngineerAndDateIPropertiesButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public SetEngineerAndDateIPropertiesButton()
		{
			DisplayName = $"Set Engineer{System.Environment.NewLine}and Date";
			InternalName = "setEngineerAndDate";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.setEngineerAndDate, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.setEngineerAndDate, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc is null) return;

			try
			{
				var designPropertySet = doc.PropertySets["Design Tracking Properties"];
				designPropertySet["Designer"].Value = SettingService.GetSetting("Engineer");
				designPropertySet["Creation Time"].Value = DateTime.Today.ToString("d");

				var summaryPropertySet = doc.PropertySets["Inventor Summary Information"];
				summaryPropertySet["Author"].Value = SettingService.GetSetting("Engineer");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.InnerException.ToString());
				return;
			}
		}
	}
}
