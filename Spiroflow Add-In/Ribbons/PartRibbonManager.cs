using Inventor;

namespace SpiroflowAddIn.Ribbons
{
	public class PartRibbonManager : BaseRibbon
	{

		public PartRibbonManager(Application inventorApp, UserInterfaceManager UIManager, string AddInGUID) : base(inventorApp, UIManager, AddInGUID)
		{
		}

		public override void CreateRibbonPanels()
		{
			ribbon = UIManager.Ribbons["Part"];
			ribbonTab = ribbon.RibbonTabs["id_TabModel"]; //must use internal names of tabs
			panels.Add(ribbonTab.RibbonPanels.Add("Spiroflow", "partPanel", AddInGUID));
		}
	}
}