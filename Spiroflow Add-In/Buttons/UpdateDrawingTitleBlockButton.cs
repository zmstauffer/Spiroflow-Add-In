using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.iLogic.Automation;
using Autodesk.iLogic.Interfaces;
using Inventor;
using SpiroflowAddIn.Utilities;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Buttons
{
	class UpdateDrawingTitleBlockButton :IButton
	{
		public Application invApp { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public string ruleText
		{
			get
			{
				string[] ruleText =
				{
					$"iLogicVb.UpdateWhenDone = True",
					$"Dim fileName As String{System.Environment.NewLine}",
					$"If ThisDrawing.ModelDocument Is Nothing Then",
					$"\t'reset drawing iproperties back to blank",
					$"\tiProperties.Value(\"Summary\", \"Title\") = \"\"",
					$"\tiProperties.Value(\"Project\", \"Part Number\") = \"\"",
					$"\tiProperties.Value(\"Project\", \"Description\") = \"\"",
					$"\tiProperties.Value(\"Project\", \"Designer\") = \"\"",
					$"\tExit Sub",
					$"End If{System.Environment.NewLine}",
					$"fileName = IO.Path.GetFileName(ThisDrawing.ModelDocument.FullFileName){System.Environment.NewLine}",
					$"On Error Resume Next",
					$"iProperties.Value(\"Summary\", \"Title\") = iProperties.Value(fileName, \"Summary\", \"Title\")",
					$"iProperties.Value(\"Project\", \"Part Number\") = iProperties.Value(fileName, \"Project\", \"Part Number\")",
					$"iProperties.Value(\"Project\", \"Description\") = iProperties.Value(fileName, \"Project\", \"Description\")",
					$"iProperties.Value(\"Project\", \"Designer\") = iProperties.Value(fileName, \"Project\", \"Designer\")",
				};
				return string.Join($"{System.Environment.NewLine}", ruleText);
			}
		}

		public UpdateDrawingTitleBlockButton()
		{
			DisplayName = $"Update{System.Environment.NewLine}Title Block";
			InternalName = "updateTitleBlock";
			PanelID = "miscPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;

			//create iLogic rule to update Title Block
			var ruleName = "UPDATE";

			var ilogicAddin = invApp.ApplicationAddIns.ItemById["{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}"];
			var automation = (iLogicAutomation)ilogicAddin.Automation;

			//add ilogic rule, but disable ilogic first. This is because some drawings have an "Update" rule that requires the model to update but fails. Also replace title block
			try
			{
				automation.RulesOnEventsEnabled = false;

				var rule = automation.GetRule(doc, ruleName);

				if (rule != null) automation.DeleteRule(doc, ruleName);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Couldn't delete existing iLogic rule, error: {ex}");
			}
			finally
			{
				automation.AddRule(doc, ruleName, ruleText);
				automation.RulesOnEventsEnabled = true;
				automation.RunRule(doc, ruleName);
			}

			//replace title block, separate try/catch block because we want to do both of these things, even if one fails
			try
			{

				var templateLocation = $@"C:\workspace\z_Documentation\TEMPLATES\SPIROFLOW MANUFACTURING - IMPERIAL.idw";
				DrawingDocument templateDoc = (DrawingDocument)invApp.Documents.Open(templateLocation, false);

				var titleBlock = templateDoc.TitleBlockDefinitions["Spiroflow Manufacturing"];

				ReplaceTitleBlock(drawingDoc, titleBlock);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Couldn't replace titleblock. Error: {ex}", "ERROR");
			}
		}

		private void ReplaceTitleBlock(DrawingDocument drawingDoc, TitleBlockDefinition titleBlockDefinition)
		{
			try
			{
				//we have to delete any titleblocks that have the same name as this one, as well as delete any titleblock currently in use
				var existingTitleBlockDef = drawingDoc.TitleBlockDefinitions[titleBlockDefinition.Name];

				if (existingTitleBlockDef != null)
				{
					foreach (Sheet sheet in drawingDoc.Sheets)
					{
						if (sheet.TitleBlock == null) continue;
						if (sheet.TitleBlock.Definition == existingTitleBlockDef) sheet.TitleBlock.Delete();
					}

					foreach (SheetFormat format in drawingDoc.SheetFormats)
					{
						if (format.ReferencedTitleBlockDefinition == existingTitleBlockDef) format.Delete();
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Couldn't delete existing titleblock or sheetformat. Error: {ex}");
			}
			finally
			{
				//we are assuming either that title block doesn't exist already, or we can safely replace it now
				titleBlockDefinition.CopyTo((_DrawingDocument)drawingDoc, true);

				var sheetTitleBlockDefinition = drawingDoc.TitleBlockDefinitions[titleBlockDefinition.Name];

				foreach (Sheet sheet in drawingDoc.Sheets)
				{
					sheet.AddTitleBlock(sheetTitleBlockDefinition);
					if (sheet.CustomTables.Count != 0)
					{
						for (int i = sheet.CustomTables.Count; i >= 0; i--)
						{
							sheet.CustomTables[i].Delete();
						}
					}
				}
				
			}
		}
	}
}
