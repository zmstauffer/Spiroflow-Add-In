﻿using System;
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
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.createPDFandDWGIcon1);
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;

			DWGPrinter.Print(drawingDoc, invApp);
			PDFPrinter.Print(drawingDoc, invApp);
		}
	}
}
