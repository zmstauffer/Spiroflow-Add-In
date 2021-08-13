using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Inventor;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	public class RenumberBOMButton
	{
		public Dictionary<int, int> itemNumberDictionary = new Dictionary<int, int>();
		private PartsList partList { get; set; }

		public void Execute() 
		{
			var doc = GetInventorApp.GetApp().ActiveDocument;

			if (doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

			DrawingDocument drawingDoc = (DrawingDocument)doc;
			Sheet currentSheet = drawingDoc.ActiveSheet;

			try
			{
				partList = currentSheet.PartsLists[1];		//YES INDEX STARTS AT 1
			}
			catch (Exception e)
			{
				return;
			}

			PartsListColumn partNumColumn = partList.PartsListColumns["PART NUMBER"];

			foreach (PartsListRow row in partList.PartsListRows)
			{
				PartsListCell cell = row[partNumColumn];
				int desiredItemNumber = 0;

				if (cell.Value.Contains("-P"))
				{
					int itemNumber = Convert.ToInt32(new Regex(@"(?:-P0*)(?<itemCode>\d+)")
						.Match(cell.Value)
						.Groups["itemCode"].Value);

					if (itemNumber != 0) desiredItemNumber = itemNumber;
				}

				if (desiredItemNumber != 0) SetItemNumber(row, desiredItemNumber, true);
				else SetItemNumber(row, GetGoodItemNumber(1), false);
			}

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
