﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows;
using Inventor;
using SpiroflowAddIn.Utilities;
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
		public ButtonDefinition buttonDef { get; set; }

		public SetEngineerAndDateIPropertiesButton()
		{
			DisplayName = $"Set Engineer{System.Environment.NewLine}and Date";
			InternalName = "setEngineerAndDate";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.test);
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc is null) return;

			try
			{
				var designPropertySet = doc.PropertySets["Design Tracking Properties"];
				designPropertySet["Designer"].Value = "zstauffer";
				designPropertySet["Creation Time"].Value = DateTime.Today.ToString("d");

				var summaryPropertySet = doc.PropertySets["Inventor Summary Information"];
				summaryPropertySet["Author"].Value = "zstauffer";
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.InnerException.ToString());
				return;
			}
		}
	}
}
