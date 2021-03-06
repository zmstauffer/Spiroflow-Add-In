using Autodesk.iLogic.Automation;
using Inventor;
using SpiroflowAddIn.Utilities;
using System;
using System.Windows;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Buttons
{
	class UpdateDrawingButton : IButton
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

		public UpdateDrawingButton()
		{
			DisplayName = $"Update{System.Environment.NewLine}Drawing";
			InternalName = "updateDrawingButton";
			PanelID = "miscPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;

			AddiLogicUpdateRule(doc);

			ReplaceTitleBlock(drawingDoc);

			UpdateStyles(drawingDoc);

			CopySketchedSymbols(drawingDoc);

			drawingDoc.Save();
		}

		private void CopySketchedSymbols(DrawingDocument drawingDoc)
		{
			var templateLocation = $@"C:\workspace\z_Documentation\TEMPLATES\SPIROFLOW MANUFACTURING - IMPERIAL.idw";

			DrawingDocument templateDoc = (DrawingDocument)invApp.Documents.Open(templateLocation, false);
			foreach (SketchedSymbolDefinition symbolDefinition in templateDoc.SketchedSymbolDefinitions)
			{
				try
				{
					var test = drawingDoc.SketchedSymbolDefinitions[symbolDefinition.Name].Name;

					//have to delete sketched symbols off all sheets before we can replace definition
					foreach (Sheet sheet in drawingDoc.Sheets)
					{
						for (int i = sheet.SketchedSymbols.Count; i > 0; i--)
						{
							if (sheet.SketchedSymbols[i].Name == test) sheet.SketchedSymbols[i].Delete();
						}
					}
					drawingDoc.SketchedSymbolDefinitions[test].Delete();
					symbolDefinition.CopyTo((_DrawingDocument)drawingDoc);
				}
				catch (Exception e)
				{
					symbolDefinition.CopyTo((_DrawingDocument)drawingDoc);
				}
			}

			templateDoc.Close();
		}

		private void UpdateStyles(DrawingDocument drawingDoc)
		{
			var styles = drawingDoc.StylesManager.Styles;

			foreach (Inventor.Style style in styles)
			{
				if (!style.UpToDate) style.UpdateFromGlobal();
			}
		}

		private void AddiLogicUpdateRule(Document doc)
		{
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

				AddEventTriggers(doc, ruleName);
			}
		}

		private void AddEventTriggers(Document doc, string ruleName)
		{
			//setup event triggers for new rule
			PropertySet iLogicPropSet = null;
			try
			{
				iLogicPropSet = doc.PropertySets["{2C540830-0723-455E-A8E2-891722EB4C3E}"];
				if (iLogicPropSet.InternalName != "{2C540830-0723-455E-A8E2-891722EB4C3E}")
				{
					iLogicPropSet.Delete();
					doc.PropertySets.Add("iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}");
					iLogicPropSet = doc.PropertySets["{2C540830-0723-455E-A8E2-891722EB4C3E}"];
				}
			}
			catch
			{
				try
				{
					doc.PropertySets.Add("iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}");
					iLogicPropSet = doc.PropertySets["{2C540830-0723-455E-A8E2-891722EB4C3E}"];
				}
				catch
				{
					MessageBox.Show("Unable to add Event Triggers");
				}
			}
			finally
			{
				if (iLogicPropSet != null)
				{
					iLogicPropSet.Add(ruleName, "AfterDocOpen", 410);       //should technically check to make sure there isn't another event trigger at 410, but I think they are sequential and have a hard time believing someone will have 10 rules on a drawing file	
					iLogicPropSet.Add(ruleName, "BeforeDocSave", 710);      //should technically check to make sure there isn't another event trigger at 710, but I think they are sequential and have a hard time believing someone will have 10 rules on a drawing file	
				}
			}
		}

		private void ReplaceTitleBlock(DrawingDocument drawingDoc)
		{
			//replace title block, separate try/catch block because we want to do both of these things, even if one fails
			var templateLocation = $@"C:\workspace\z_Documentation\TEMPLATES\SPIROFLOW MANUFACTURING - IMPERIAL.idw";

			DrawingDocument templateDoc = (DrawingDocument)invApp.Documents.Open(templateLocation, false);

			TitleBlockDefinition titleBlockDefinition = null;

			if (templateDoc == null) return;

			try
			{
				titleBlockDefinition = templateDoc.TitleBlockDefinitions["Spiroflow Manufacturing"];

			}
			catch (Exception ex)
			{
				MessageBox.Show($"Couldn't replace title block. Error: {ex}", "ERROR");
				templateDoc.Close();
				return;
			}

			if (titleBlockDefinition == null) return;

			try
			{
				//we have to delete any title blocks in use
				for (int i = drawingDoc.Sheets.Count; i > 0; i--)
				{
					if (drawingDoc.Sheets[i].TitleBlock != null) drawingDoc.Sheets[i].TitleBlock.Delete();
				}
			}
			catch (Exception ex)
			{
				
			}
			finally
			{
				//we are assuming either that title block doesn't exist already, or we can add a new one that isn't replacing the old one
				var newTitleBlockDef = titleBlockDefinition.CopyTo((_DrawingDocument)drawingDoc, false);

				var sheetTitleBlockDefinition = drawingDoc.TitleBlockDefinitions[newTitleBlockDef.Name];

				foreach (Sheet sheet in drawingDoc.Sheets)
				{
					try
					{
						if (sheet.Name.ToUpper().Contains("FLAT")) continue;        //skip sheets that have flat patterns

						sheet.AddTitleBlock(sheetTitleBlockDefinition);

						if (sheet.CustomTables.Count != 0)
						{
							for (int i = sheet.CustomTables.Count; i > 0; i--)
							{
								sheet.CustomTables[i].Delete();
							}
						}
					}
					catch
					{
						continue;
					}
				}

				templateDoc.Close();
			}
		}
	}
}
