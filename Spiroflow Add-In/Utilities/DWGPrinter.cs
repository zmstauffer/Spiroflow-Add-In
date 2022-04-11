using Inventor;
using System.Windows;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Utilities
{
	static class DWGPrinter
	{
		public static void Print(DrawingDocument drawingDoc, Application invApp)
		{
			//make dwg on a per-sheet basis
			TranslatorAddIn DWGAddin = (TranslatorAddIn)invApp.ApplicationAddIns.ItemById["{C24E3AC2-122E-11D5-8E91-0010B541CD80}"];
			TranslationContext context = invApp.TransientObjects.CreateTranslationContext();
			NameValueMap options = invApp.TransientObjects.CreateNameValueMap();
			DataMedium dataMedium = invApp.TransientObjects.CreateDataMedium();

			string iniFilename = @"C:\workspace\AutoCAD export settings.ini";

			if (!System.IO.File.Exists(@"C:\workspace\AutoCAD export settings.ini"))
			{
				MessageBox.Show($@"Cannot find {iniFilename}, please generate AutoCAD export settings and save them to {iniFilename} for DWG export to work.", "ERROR");
				return;
			}

			string filepath = SettingService.GetSetting("DrawingExportPath");

			context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

			if (DWGAddin.HasSaveCopyAsOptions[drawingDoc, context, options]) options.Value["Export_Acad_IniFile"] = iniFilename;

			if (!System.IO.Directory.Exists(filepath)) System.IO.Directory.CreateDirectory(filepath);

			int sheetNum = 1;
			int totalSheets = drawingDoc.Sheets.Count;
			var revision = drawingDoc.PropertySets["Inventor Summary Information"]["Revision Number"].Value;
			string drawingName = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName);

			var revisionText = revision == "1" ? "" : $" rev {revision}";

			foreach (Sheet sheet in drawingDoc.Sheets)
			{
				sheet.Activate();
				sheet.Update();

				if (drawingDoc.Sheets.Count > 1) dataMedium.FileName = $"{filepath}{drawingName}{revisionText} Sheet {sheetNum} of {totalSheets}.dwg";
				else dataMedium.FileName = $"{filepath}{drawingName}{revisionText}.dwg";

				DWGAddin.SaveCopyAs(drawingDoc, context, options, dataMedium);

				sheetNum++;
			}
		}
	}
}
