using Inventor;

namespace SpiroflowAddIn.Utilities
{
	static class PDFPrinter
	{
		public static void Print(DrawingDocument drawingDoc, Application invApp)
		{
			TranslatorAddIn PDFAddin = (TranslatorAddIn)invApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"];
			TranslationContext context = invApp.TransientObjects.CreateTranslationContext();
			NameValueMap options = invApp.TransientObjects.CreateNameValueMap();
			DataMedium dataMedium = invApp.TransientObjects.CreateDataMedium();

			string filepath = SettingService.GetSetting("DrawingExportPath");
			string fileName = System.IO.Path.GetFileNameWithoutExtension(drawingDoc.DisplayName);
			var revision = drawingDoc.PropertySets["Inventor Summary Information"]["Revision Number"].Value;
			var revisionText = revision == "1" ? "" : $" rev {revision}";

			context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

			if (!PDFAddin.Activated) PDFAddin.Activate();

			options.Value["All_Color_AS_Black"] = 0;
			options.Value["Remove_Line_Weights"] = 0;
			options.Value["Vector_Resolution"] = 400;
			options.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;

			dataMedium.FileName = $"{filepath}{fileName}{revisionText}.pdf";

			foreach (Sheet sheet in drawingDoc.Sheets)
			{
				sheet.Activate();
				sheet.Update();
				foreach (DrawingView view in sheet.DrawingViews)
				{
					do
					{
						invApp.UserInterfaceManager.DoEvents();
						invApp.StatusBarText = "Updating drawing views";
					} while (view.IsUpdateComplete == false);
				}
			}

			PDFAddin.SaveCopyAs(drawingDoc, context, options, dataMedium);
		}
	}
}
