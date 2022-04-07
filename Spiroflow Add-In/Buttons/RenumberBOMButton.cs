using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Inventor;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	public class RenumberBOMButton : IButton
	{
		public Application invApp { get; set; }
		public string DisplayName {get;set;}
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		public Dictionary<int, int> itemNumberDictionary { get; set; }
		private PartsList partList { get; set; }

		public RenumberBOMButton()
		{
			DisplayName = $"Renumber{System.Environment.NewLine}BOM";
			InternalName = "renumberBOM";
			PanelID = "bomPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.renumberBOM, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.renumberBOM, 16, 16));
		}

		public void Execute(NameValueMap context) 
		{
			var doc = invApp.ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			itemNumberDictionary = new Dictionary<int, int>();
			DrawingDocument drawingDoc = (DrawingDocument)doc;
			Sheet currentSheet = drawingDoc.ActiveSheet;

			try
			{
				partList = currentSheet.PartsLists[1];		//YES INDEX STARTS AT 1
			}
			catch (Exception e)
			{
				Debug.WriteLine(e.Message);
				return;
			}

			PartsListColumn partNumColumn = partList.PartsListColumns["PART NUMBER"];

			foreach (PartsListRow row in partList.PartsListRows)
			{
				PartsListCell cell = row[partNumColumn];
				int desiredItemNumber = 0;

				if (cell.Value.Contains("-P"))
				{
					int itemNumber = 0;
					try
					{
						itemNumber = Convert.ToInt32(new Regex(@"(?:-P0*)(?<itemCode>\d+)")
						  .Match(cell.Value)
						  .Groups["itemCode"].Value);
					}
					catch
					{ }

					if (itemNumber != 0) desiredItemNumber = itemNumber;
				}

				if (desiredItemNumber != 0) SetItemNumber(row, desiredItemNumber, true);
				else SetItemNumber(row, GetGoodItemNumber(1), false);
			}

			//sort and save BOM overrides to model
			partList.Sort2("ITEM", AutoSortOnUpdate: true);
			partList.SaveItemOverridesToBOM();
			return;
		}

		private int GetGoodItemNumber(int badNumber)
		{
			int goodNumber = badNumber;
			bool badEntry = true;

			while (badEntry)
			{
				if (itemNumberDictionary.ContainsValue(goodNumber)) goodNumber++;
				else badEntry = false;
			}

			return goodNumber;
		}

		private void SetItemNumber(PartsListRow row, int desiredItemNumber, bool forceItemNumber)
		{
			PartsListColumn col = partList.PartsListColumns["ITEM"];

			if (col is null) return;

			if (itemNumberDictionary.ContainsKey(desiredItemNumber))
			{
				if (forceItemNumber) ChangeItemNumber(desiredItemNumber);
				else desiredItemNumber = GetGoodItemNumber(itemNumberDictionary[desiredItemNumber]);
			}

			PartsListCell cell = row[col];
			cell.Value = desiredItemNumber.ToString();
			itemNumberDictionary.Add(desiredItemNumber, desiredItemNumber);
		}

		private void ChangeItemNumber(int numToChange)
		{
			int goodNum = GetGoodItemNumber(numToChange);
			PartsListRow row = partList.PartsListRows[itemNumberDictionary[numToChange]];
			PartsListColumn col = partList.PartsListColumns["ITEM"];
			row[col].Value = goodNum.ToString();

			itemNumberDictionary.Remove(numToChange);
			itemNumberDictionary.Add(goodNum, goodNum);
		}
	}
}
