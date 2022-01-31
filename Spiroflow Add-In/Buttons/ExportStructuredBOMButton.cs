using Inventor;
using SpiroflowAddIn.Utilities;
using ClosedXML.Excel;
using Microsoft.VisualBasic.Compatibility.VB6;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml.Drawing;
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
		public ButtonDefinition buttonDef { get; set; }
		public IXLWorkbook workbook { get; set; }
		public IXLWorksheet worksheet { get; set; }

		private BOM assyBOM { get; set; }
		private const string structuredBOMImportFilename = @"C:\workspace\bom - structured view.xml";
		private const string partslistBOMImportFilename = @"C:\workspace\bom - partslist view.xml";

		public ExportStructuredBOMButton()
		{
			DisplayName = $"Export BOM";
			InternalName = "exportBOM";
			PanelID = "assemblyPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(Properties.Resources.billOfMaterialsIcon);
		}

		public void Execute(NameValueMap context)
		{
			if (invApp.ActiveDocument.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assyDoc = (AssemblyDocument)invApp.ActiveDocument;
			AssemblyComponentDefinition assyDef = assyDoc.ComponentDefinition;
			assyBOM = assyDef.BOM;

			if (assyBOM is null) return;

			assyBOM.PartsOnlyViewEnabled = true;
			assyBOM.StructuredViewFirstLevelOnly = false;
			assyBOM.StructuredViewEnabled = true;

			string templateFilename = @"Y:\BOM\HEADER TEMPLATE.xlsx";

			using (workbook = new XLWorkbook(templateFilename))
			{
				ExportStructuredBOM(workbook);
				ExportPartsListBOM(workbook);
				string newFilename = System.IO.Path.GetFileNameWithoutExtension(assyDoc.DisplayName);
				workbook.SaveAs(@"C:\workspace\" + newFilename + "-BOM.xlsx");
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

			for (int i = 1; i <= 10; i++)
			{
				worksheet.Column(i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
			}

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

			for (int i = 1; i <= 10; i++)
			{
				worksheet.Column(i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
			}

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
	}
}
