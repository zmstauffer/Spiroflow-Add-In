using ClosedXML.Excel;
using Inventor;
using Microsoft.VisualBasic.Compatibility.VB6;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Application = Inventor.Application;

namespace SpiroflowAddIn.Buttons
{
	public class ExportStructuredBOMButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }
		public IXLWorkbook workbook { get; set; }
		public IXLWorksheet worksheet { get; set; }
		private BOM assyBOM { get; set; }
		private string structuredBOMImportFilename = $@"{SettingService.GetSetting("ConfigFilesPath")}bom - structured view.xml";
		private string partslistBOMImportFilename = $@"{SettingService.GetSetting("ConfigFilesPath")}bom - partslist view.xml";

		public ExportStructuredBOMButton()
		{
			DisplayName = $"Export BOM";
			InternalName = "exportBOM";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.billOfMaterialsIcon, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.billOfMaterialsIcon, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			if (invApp.ActiveDocument.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assyDoc = (AssemblyDocument)invApp.ActiveDocument;
			AssemblyComponentDefinition assyDef = assyDoc.ComponentDefinition;

			if (assyDef.Occurrences.Count <= 0) return;

			assyBOM = assyDef.BOM;

			if (assyBOM is null) return;

			assyBOM.PartsOnlyViewEnabled = true;
			assyBOM.StructuredViewFirstLevelOnly = false;
			assyBOM.StructuredViewEnabled = true;

			string configFilesPath = SettingService.GetSetting("ConfigFilesPath");
			string templateFilename = $"{configFilesPath}HEADER TEMPLATE.xlsx";
			structuredBOMImportFilename = $"{configFilesPath}bom - structured view.xml";
			partslistBOMImportFilename = $"{configFilesPath}bom - partslist view.xml";

			if (!VaultFunctions.CheckFileAndDownloadIfNecessary(structuredBOMImportFilename)) return;
			if (!VaultFunctions.CheckFileAndDownloadIfNecessary(partslistBOMImportFilename)) return;
			if (!VaultFunctions.CheckFileAndDownloadIfNecessary(templateFilename)) return;

			using (workbook = new XLWorkbook(templateFilename))
			{
				ExportStructuredBOM(workbook);
				ExportPartsListBOM(workbook);
				FillOutHeader(workbook);

				string newFilename = System.IO.Path.GetFileNameWithoutExtension(assyDoc.DisplayName);
				string outputPath = SettingService.GetSetting("BOMExportPath");

				workbook.SaveAs($@"{outputPath}{newFilename}-BOM.xlsx");
			}
		}

		private void ExportStructuredBOM(IXLWorkbook workbook)
		{
			worksheet = workbook.Worksheet("Subassembly BOM");

			assyBOM.StructuredViewDelimiter = ".";
			assyBOM.ImportBOMCustomization(structuredBOMImportFilename);

			BOMView bomView = assyBOM.BOMViews["Structured"];
			bomView.Sort("Part Number", true);
			bomView.Renumber();

			int startRow = 13;
			int currentRow = 14;

			//C2 is date. C3 is Rev. C4 is "BY". E2 is Model #. E3 is Job#/Serial#. E4 is customer. G3 is PO#.
			//retrieve header data from model
			var date = DateTime.Today.ToShortDateString();
			string designer = IPropService.GetParameter<string>("Designer");
			string projectNumber = IPropService.GetParameter<string>("Project");
			string purchaseOrder = IPropService.GetParameter<string>("Subject");
			string company = IPropService.GetParameter<string>("Company");
			string description = IPropService.GetParameter<string>("Description");

			worksheet.Range("C2").Value = date;
			worksheet.Range("C4").Value = designer;
			worksheet.Range("E2").Value = description;
			worksheet.Range("E3").Value = projectNumber;
			worksheet.Range("E4").Value = company;
			worksheet.Range("G3").Value = purchaseOrder;

			try
			{
				foreach (BOMRow bomViewRow in bomView.BOMRows)
				{
					if (bomViewRow.ComponentDefinitions[1].Type == ObjectTypeEnum.kVirtualComponentDefinitionObject)
					{
						OutputVirtualComponentToExcel(currentRow, bomViewRow);
					}
					else
					{
						OutputToExcelRow(currentRow, bomViewRow);
					}

					if (bomViewRow.ChildRows != null & bomViewRow.BOMStructure != BOMStructureEnum.kInseparableBOMStructure)
					{
						currentRow = OutputChildren(bomViewRow, currentRow);
					}

					currentRow += 1;
				}
			}
			catch (Exception ex)
			{
				Console.Write(ex.Message);
			}
		}

		private void ExportPartsListBOM(IXLWorkbook workbook)
		{
			worksheet = workbook.Worksheet("Purchasing BOM");

			//assyBOM.StructuredViewDelimiter = ".";
			assyBOM.ImportBOMCustomization(partslistBOMImportFilename);

			BOMView bomView = assyBOM.BOMViews["Parts Only"];
			bomView.Sort("Part Number", true);
			bomView.Renumber();

			int startRow = 13;
			int currentRow = 14;

			worksheet.Range("A" + startRow).Value = "Item";
			worksheet.Range("B" + startRow).Value = "QTY";
			worksheet.Range("C" + startRow).Value = "Part Number";
			worksheet.Range("D" + startRow).Value = "Thumbnail";
			worksheet.Range("E" + startRow).Value = "Description";
			worksheet.Range("F" + startRow).Value = "Material";
			worksheet.Range("G" + startRow).Value = "Vendor";

			worksheet.Column(1).Width = 5;
			worksheet.Column(2).Width = 8.5d;
			worksheet.Column(3).Width = 19;
			worksheet.Column(4).Width = 9.3d;
			worksheet.Column(5).Width = 60;
			worksheet.Column(6).Width = 19;
			worksheet.Column(7).Width = 20;
			worksheet.Column(8).Width = 13;
			worksheet.Column(9).Width = 13;

			//B2 is date. B3 is Rev. B4 is "BY". D2 is Model #. D3 is Job#/Serial#. D4 is customer. F3 is PO#.
			//retrieve header data from model
			var date = DateTime.Today.ToShortDateString();
			string designer = IPropService.GetParameter<string>("Designer");
			string projectNumber = IPropService.GetParameter<string>("Project");
			string purchaseOrder = IPropService.GetParameter<string>("Subject");
			string company = IPropService.GetParameter<string>("Company");
			string description = IPropService.GetParameter<string>("Description");

			worksheet.Range("B2").Value = date;
			worksheet.Range("B4").Value = designer;
			worksheet.Range("D2").Value = description;
			worksheet.Range("D3").Value = projectNumber;
			worksheet.Range("D4").Value = company;
			worksheet.Range("F3").Value = purchaseOrder;

			try
			{
				foreach (BOMRow bomViewRow in bomView.BOMRows)
				{
					if (bomViewRow.ComponentDefinitions[1].Type == ObjectTypeEnum.kVirtualComponentDefinitionObject)
					{
						OutputPartVirtualComponentToExcel(currentRow, bomViewRow);
					}
					else
					{
						OutputPartToExcelRow(currentRow, bomViewRow);
					}

					currentRow += 1;
				}
			}
			catch (Exception ex)
			{
				Console.Write(ex.Message);
			}
		}

		public int OutputChildren(BOMRow bomRow, int counter)
		{
			List<BOMRow> rowList = new List<BOMRow>();

			foreach (BOMRow childBOMRow in bomRow.ChildRows)
			{
				rowList.Add(childBOMRow);
			}

			rowList = rowList.OrderBy(x => x.ItemNumber.Replace(".", "")).ToList();

			foreach (BOMRow childBOMRow in rowList)
			{
				counter += 1;
				if (childBOMRow.ComponentDefinitions[1].Type == ObjectTypeEnum.kVirtualComponentDefinitionObject)
				{
					OutputVirtualComponentToExcel(counter, childBOMRow, true);
				}
				else
				{
					OutputToExcelRow(counter, childBOMRow, true);
				}

				if (childBOMRow.ChildRows != null && childBOMRow.BOMStructure != BOMStructureEnum.kInseparableBOMStructure) counter = OutputChildren(childBOMRow, counter);
			}

			return counter;
		}

		public void OutputToExcelRow(int rowNum, BOMRow currentBOMRow, bool subItem = false)
		{
			string picFilename = @"C:\workspace\tempPic.jpg";

			Document subDoc = (Document)currentBOMRow.ComponentDefinitions[1].Document;
			PropertySet subDocDesignPropertySet = subDoc.PropertySets["Design Tracking Properties"];
			PropertySet subDocSummaryInfo = subDoc.PropertySets["Inventor Summary Information"];
			PropertySet subDocDocInfo = subDoc.PropertySets["Document Summary Information"];

			bool boldText = currentBOMRow.ChildRows != null & currentBOMRow.BOMStructure != BOMStructureEnum.kInseparableBOMStructure;

			if (subItem)
			{
				worksheet.Range("B" + rowNum).Style.NumberFormat.Format = "@";
				worksheet.Range("B" + rowNum).Value = currentBOMRow.ItemNumber;
			}
			else
			{
				worksheet.Range("A" + rowNum).Value = currentBOMRow.ItemNumber;
			}

			if (subItem)
			{
				var quantity = GetTotalQuantity(currentBOMRow);
				worksheet.Range("C" + rowNum).Value = quantity;
			}
			else worksheet.Range("C" + rowNum).Value = currentBOMRow.TotalQuantity;

			worksheet.Range("D" + rowNum).Value = subDocDesignPropertySet["Part Number"].Value;
			worksheet.Range("D" + rowNum).Style.Alignment.WrapText = true;
			worksheet.Range("F" + rowNum).Value = subDocDesignPropertySet["Description"].Value;

			// need to get weld material if weldment
			string material = "";
			material = (string)subDocDesignPropertySet["Material"].Value;
			if (string.IsNullOrEmpty(material))
			{
				material = (string)subDocDesignPropertySet["Weld Material"].Value;
			}

			worksheet.Range("G" + rowNum).Value = material;
			worksheet.Range("H" + rowNum).Value = subDocDesignPropertySet["Vendor"].Value;

			// this section generates the thumbnail and puts it into excel file
			// make temp picture
			var pic = Support.IPictureDispToImage(subDoc.Thumbnail);
			pic.Save(picFilename);

			// put the picture in the excel file and format it
			var picToInsert = worksheet.AddPicture(picFilename).MoveTo(worksheet.Cell(rowNum, 5));
			picToInsert.Width = 70;
			picToInsert.Height = 70;

			// set row height to fit picture
			worksheet.Row(rowNum).Height = 53;

			//set text to bold
			if (boldText)
			{
				worksheet.Row(rowNum).Style.Font.Bold = true;
				worksheet.Row(rowNum).Style.Font.FontSize = 12;
			}
		}

		public void OutputVirtualComponentToExcel(int rowNum, BOMRow currentBOMRow, bool subItem = false)
		{
			VirtualComponentDefinition subDefinition = (VirtualComponentDefinition)currentBOMRow.ComponentDefinitions[1];
			PropertySet subDocDesignPropertySet = subDefinition.PropertySets["Design Tracking Properties"];
			PropertySet subDocSummaryInfo = subDefinition.PropertySets["Inventor Summary Information"];
			PropertySet subDocDocInfo = subDefinition.PropertySets["Document Summary Information"];

			bool boldText = currentBOMRow.ChildRows != null & currentBOMRow.BOMStructure != BOMStructureEnum.kInseparableBOMStructure;

			if (subItem)
			{
				worksheet.Range("B" + rowNum).Style.NumberFormat.Format = "@";
				worksheet.Range("B" + rowNum).Value = currentBOMRow.ItemNumber;
			}
			else
			{
				worksheet.Range("A" + rowNum).Value = currentBOMRow.ItemNumber;
			}

			if (subItem)
			{
				var quantity = GetTotalQuantity(currentBOMRow);
				worksheet.Range("C" + rowNum).Value = quantity;
			}
			else worksheet.Range("C" + rowNum).Value = currentBOMRow.TotalQuantity;

			worksheet.Range("D" + rowNum).Value = subDocDesignPropertySet["Part Number"].Value;
			worksheet.Range("D" + rowNum).Style.Alignment.WrapText = true;
			worksheet.Range("F" + rowNum).Value = subDocDesignPropertySet["Description"].Value;

			// need to get weld material if weldment
			string material = "";
			material = (string)subDocDesignPropertySet["Material"].Value;
			if (string.IsNullOrEmpty(material))
			{
				material = (string)subDocDesignPropertySet["Weld Material"].Value;
			}

			worksheet.Range("G" + rowNum).Value = material;
			worksheet.Range("H" + rowNum).Value = subDocDesignPropertySet["Vendor"].Value;

			// set row height to fit picture
			worksheet.Row(rowNum).Height = 53;

			//set text to bold
			if (boldText)
			{
				worksheet.Row(rowNum).Style.Font.Bold = true;
				worksheet.Row(rowNum).Style.Font.FontSize = 12;
			}
		}

		private string GetTotalQuantity(BOMRow bomRow)
		{
			var qtyString = bomRow.TotalQuantity;
			try
			{
				var numberPart = Regex.Match(qtyString, "[0-9.]*");
				var unitString = Regex.Match(qtyString, "[A-z]+");

				if (!numberPart.Success)
				{
					MessageBox.Show($"Can't determine quantity of {bomRow.ItemNumber}");
					return bomRow.TotalQuantity;
				}

				BOMRow parentRow = (BOMRow)bomRow.Parent;

				string totalQty = (Convert.ToDouble(numberPart.Value) * Convert.ToDouble(parentRow.TotalQuantity)).ToString();

				return $"{totalQty} {unitString.Value}";
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Can't determine quantity of {bomRow.ItemNumber}");
				Console.Write(ex.Message);
				return bomRow.TotalQuantity;
			}
		}

		public void OutputPartToExcelRow(int rowNum, BOMRow currentBOMRow)
		{
			string picFilename = @"C:\workspace\tempPic.jpg";

			Document subDoc = (Document)currentBOMRow.ComponentDefinitions[1].Document;
			PropertySet subDocDesignPropertySet = subDoc.PropertySets["Design Tracking Properties"];
			PropertySet subDocSummaryInfo = subDoc.PropertySets["Inventor Summary Information"];
			PropertySet subDocDocInfo = subDoc.PropertySets["Document Summary Information"];

			worksheet.Range("A" + rowNum).Value = currentBOMRow.ItemNumber;
			worksheet.Range("B" + rowNum).Value = currentBOMRow.TotalQuantity;
			worksheet.Range("C" + rowNum).Value = subDocDesignPropertySet["Part Number"].Value;
			worksheet.Range("C" + rowNum).Style.Alignment.WrapText = true;
			worksheet.Range("E" + rowNum).Value = subDocDesignPropertySet["Description"].Value;

			// need to get weld material if weldment
			string material = "";
			material = (string)subDocDesignPropertySet["Material"].Value;
			if (string.IsNullOrEmpty(material))
			{
				material = (string)subDocDesignPropertySet["Weld Material"].Value;
			}

			worksheet.Range("F" + rowNum).Value = material;
			worksheet.Range("G" + rowNum).Value = subDocDesignPropertySet["Vendor"].Value;

			// this section generates the thumbnail and puts it into excel file
			// make temp picture
			var pic = Support.IPictureDispToImage(subDoc.Thumbnail);
			pic.Save(picFilename);

			// put the picture in the excel file and format it
			var picToInsert = worksheet.AddPicture(picFilename).MoveTo(worksheet.Cell(rowNum, 4));
			picToInsert.Width = 70;
			picToInsert.Height = 70;

			// set row height to fit picture
			worksheet.Row(rowNum).Height = 53;
		}

		public void OutputPartVirtualComponentToExcel(int rowNum, BOMRow currentBOMRow)
		{
			VirtualComponentDefinition subDefinition = (VirtualComponentDefinition)currentBOMRow.ComponentDefinitions[1];
			PropertySet subDocDesignPropertySet = subDefinition.PropertySets["Design Tracking Properties"];
			PropertySet subDocSummaryInfo = subDefinition.PropertySets["Inventor Summary Information"];
			PropertySet subDocDocInfo = subDefinition.PropertySets["Document Summary Information"];

			worksheet.Range("A" + rowNum).Value = currentBOMRow.ItemNumber;
			worksheet.Range("B" + rowNum).Value = currentBOMRow.TotalQuantity;

			worksheet.Range("C" + rowNum).Value = subDocDesignPropertySet["Part Number"].Value;
			worksheet.Range("C" + rowNum).Style.Alignment.WrapText = true;
			worksheet.Range("E" + rowNum).Value = subDocDesignPropertySet["Description"].Value;

			// need to get weld material if weldment
			string material = "";
			material = (string)subDocDesignPropertySet["Material"].Value;
			if (string.IsNullOrEmpty(material))
			{
				material = (string)subDocDesignPropertySet["Weld Material"].Value;
			}

			worksheet.Range("F" + rowNum).Value = material;
			worksheet.Range("G" + rowNum).Value = subDocDesignPropertySet["Vendor"].Value;

			// set row height to fit picture
			worksheet.Row(rowNum).Height = 53;
		}

		private void FillOutHeader(IXLWorkbook workbook)
		{
			worksheet = workbook.Worksheet("Subassembly BOM");

			assyBOM.StructuredViewDelimiter = ".";
			assyBOM.ImportBOMCustomization(structuredBOMImportFilename);

			BOMView bomView = assyBOM.BOMViews["Structured"];
			bomView.Sort("Part Number", true);
			bomView.Renumber();

			int startRow = 13;
			int currentRow = 14;

			worksheet.Range("A" + startRow).Value = "Item";
			worksheet.Range("B" + startRow).Value = "Sub Item";
			worksheet.Range("C" + startRow).Value = "QTY";
			worksheet.Range("D" + startRow).Value = "Part Number";
			worksheet.Range("E" + startRow).Value = "Thumbnail";
			worksheet.Range("F" + startRow).Value = "Description";
			worksheet.Range("G" + startRow).Value = "Material";
			worksheet.Range("H" + startRow).Value = "Vendor";

			worksheet.Column(1).Width = 5;
			worksheet.Column(2).Width = 6;
			worksheet.Column(3).Width = 8.5d;
			worksheet.Column(4).Width = 19;
			worksheet.Column(5).Width = 9.3d;
			worksheet.Column(6).Width = 60;
			worksheet.Column(7).Width = 19;
			worksheet.Column(8).Width = 20;
			worksheet.Column(9).Width = 13;
			worksheet.Column(10).Width = 13;
		}
	}
}
