using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;

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
			string filepath = @"C:\workspace\";

			context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

			if (DWGAddin.HasSaveCopyAsOptions[drawingDoc, context, options]) options.Value["Export_Acad_IniFile"] = iniFilename;

			if (!System.IO.Directory.Exists(filepath)) System.IO.Directory.CreateDirectory(filepath);

			int sheetNum = 1;
			int totalSheets = drawingDoc.Sheets.Count;
			string drawingName = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName);

			foreach (Sheet sheet in drawingDoc.Sheets)
			{
				sheet.Activate();

				if (drawingDoc.Sheets.Count > 1) dataMedium.FileName = $"{filepath}{drawingName} Sheet {sheetNum} of {totalSheets}.dwg";
				else dataMedium.FileName = $"{filepath}{drawingName}.dwg";

				DWGAddin.SaveCopyAs(drawingDoc, context, options, dataMedium);

				sheetNum++;
			}
		}
	}
}
