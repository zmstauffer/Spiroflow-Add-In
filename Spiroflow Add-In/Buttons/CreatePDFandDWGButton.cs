using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	public class CreatePDFandDWGButton : IButton
	{
		public Application invApp { get; set; }
		public string PanelID { get;set; }
		public stdole.IPictureDisp icon { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public CreatePDFandDWGButton()
		{
			DisplayName = $"Create PDF{System.Environment.NewLine}and DWGs";
			InternalName = "createPDFandDWG";
			PanelID = "printPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.createPDFandDWGIcon);
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;

			CreateDWG(drawingDoc);
			CreatePDF(drawingDoc);
		}

		private void CreateDWG(DrawingDocument drawingDoc)
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

		private void CreatePDF(DrawingDocument drawingDoc)
		{
			TranslatorAddIn PDFAddin = (TranslatorAddIn)invApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
			TranslationContext context = invApp.TransientObjects.CreateTranslationContext();
			NameValueMap options = invApp.TransientObjects.CreateNameValueMap();
			DataMedium dataMedium = invApp.TransientObjects.CreateDataMedium();

			string filepath = @"C:\workspace\";
			string fileName = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName);

			context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

			if (!PDFAddin.Activated) PDFAddin.Activate();

			options.Value["All_Color_AS_Black"] = 0;
			options.Value["Remove_Line_Weights"] = 0;
			options.Value["Vector_Resolution"] = 400;
			options.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;

			dataMedium.FileName = $"{filepath}{fileName}.pdf";

			PDFAddin.SaveCopyAs(drawingDoc, context, options, dataMedium);
		}
	}
}
