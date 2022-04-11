using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using SpiroflowAddIn.Utilities;
using Application = Inventor.Application;
using MessageBox = System.Windows.MessageBox;

namespace SpiroflowAddIn.Buttons
{
	public class PartDetailCheckButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public PartDetailCheckButton()
		{
			DisplayName = $"Check Drawing{System.Environment.NewLine}For Details";
			InternalName = "partDetailCheck";
			PanelID = "miscPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.partDetailCheck, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.partDetailCheck, 16, 16));
		}

		/// <summary>
		/// This compares the parts list to detail views on the drawing and makes sure they match
		/// </summary>
		/// <param name="context"></param>
		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;
			var activeSheet = drawingDoc.ActiveSheet;
			PartsList partsList;

			try
			{
				partsList = activeSheet.PartsLists[1];
			}
			catch (Exception ex)
			{
				MessageBox.Show("No Parts List found.", "ERROR");
				return;
			}

			List<string> partNumbers = new List<string>();
			List<string> extraItems = new List<string>();

			foreach (PartsListRow row in partsList.PartsListRows)
			{
				partNumbers.Add(row["PART NUMBER"].Value);
			}

			foreach (Sheet sheet in drawingDoc.Sheets)
			{
				foreach (DrawingView view in sheet.DrawingViews)
				{
					if (view.ShowLabel)
					{
						var firstSpaceIndex = view.Label.Text.IndexOf(" ");
						var viewPartNumber = view.Label.Text.Substring(0, firstSpaceIndex);
						
						if (partNumbers.All(x => x != viewPartNumber) && view.Label.Text.Contains("DETAIL")) extraItems.Add($"Extra View: {view.Label.Text} on Sht. {sheet.Name}");

						var partNumber = partNumbers.FirstOrDefault(x => x == viewPartNumber);
						partNumbers.Remove(partNumber);
					}
				}
			}

			foreach (string partNumber in partNumbers)
			{
				extraItems.Add($"Missing View: {partNumber}");
			}

			if (extraItems.Count > 0) MessageBox.Show(String.Join("\n", extraItems.ToArray()), "VIEW ERRORS");
			else MessageBox.Show("No extra or missing views found.", "What a Winner!");
		}
	}
}
