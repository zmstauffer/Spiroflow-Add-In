using Inventor;

namespace SpiroflowAddIn.Ribbons
{
	public class ZeroDocRibbonManager : BaseRibbon
	{

		public ZeroDocRibbonManager(Application inventorApp, UserInterfaceManager UIManager, string AddInGUID) : base(inventorApp, UIManager, AddInGUID)
		{
		}

		public override void CreateRibbonPanels()
		{
			ribbon = UIManager.Ribbons["ZeroDoc"];
			ribbonTab = ribbon.RibbonTabs.Add("Spiroflow", "id_Spiroflow_ZeroDoc", AddInGUID);
			panels.Add(ribbonTab.RibbonPanels.Add("Spiroflow", "spiroflowPanel", AddInGUID));
		}
	}
}
