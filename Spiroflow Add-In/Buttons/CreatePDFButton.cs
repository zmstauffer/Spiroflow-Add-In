using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Inventor;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	public class CreatePDFButton : IButton
	{
		public Application invApp { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public CreatePDFButton()
		{
			DisplayName = $"Create PDF";
			InternalName = "createPDF";
			PanelID = "printPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.CreatePDFButtonIcon);
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;

			PDFPrinter.Print(drawingDoc, invApp);
		}
	}
}
