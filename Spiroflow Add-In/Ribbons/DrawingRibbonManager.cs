using Inventor;
using SpiroflowAddIn.Buttons;
using SpiroflowAddIn.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Resources;

namespace SpiroflowAddIn.Ribbons
{
	class DrawingRibbonManager : BaseRibbon
	{

		public DrawingRibbonManager(Application inventorApp, UserInterfaceManager UIManager, string AddInGUID) : base(inventorApp, UIManager, AddInGUID)
		{
		}

		public override void CreateRibbonPanels()
		{
			ribbon = UIManager.Ribbons["Drawing"];
			ribbonTab = ribbon.RibbonTabs["id_TabPlaceViews"];																		//ribbon.RibbonTabs.Add("Spiroflow", "id_Spiroflow_Drawing", AddInGUID);
			panels.Add(ribbonTab.RibbonPanels.Add("BOM Functions", "bomPanel", AddInGUID));
			panels.Add(ribbonTab.RibbonPanels.Add("Miscellaneous Functions", "miscPanel", AddInGUID));
			panels.Add(ribbonTab.RibbonPanels.Add("Printing", "printPanel", AddInGUID));
		}
	}
}
