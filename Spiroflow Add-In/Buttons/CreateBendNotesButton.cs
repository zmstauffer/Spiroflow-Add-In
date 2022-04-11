using Inventor;
using SpiroflowAddIn.Utilities;
using System.Linq;

namespace SpiroflowAddIn.Buttons
{
	public class CreateBendNotesButton : IButton
	{
		public Application invApp { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public CreateBendNotesButton()
		{
			DisplayName = $"Create Bend{System.Environment.NewLine}Notes";
			InternalName = "createBendNotes";
			PanelID = "miscPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 16, 16));
		}

		/// <summary>
		/// This function deletes any existing bend notes on flat pattern drawings, then places new bend notes at all bends on all views.
		/// </summary>
		/// <param name="context"></param>
		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;
			var sheet = drawingDoc.ActiveSheet;

			//delete all bend notes so we don't duplicate them
			for (int i = sheet.DrawingNotes.BendNotes.Count; i > 0; i--)
			{
				sheet.DrawingNotes.BendNotes[i].Delete();
			}

			foreach (DrawingView drawingView in sheet.DrawingViews)
			{
				var bendCurves = drawingView.DrawingCurves.Cast<DrawingCurve>()
					.Where(x => x.EdgeType == DrawingEdgeTypeEnum.kBendDownEdge || x.EdgeType == DrawingEdgeTypeEnum.kBendUpEdge).ToList();

				var bendNotes = sheet.DrawingNotes.BendNotes;

				foreach (DrawingCurve curve in bendCurves)
				{
					bendNotes.Add(curve);
				}
			}


		}
	}
}
