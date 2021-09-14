using Inventor;

namespace SpiroflowAddIn.Ribbons
{
	public class AssemblyRibbonManager : BaseRibbon
	{

		public AssemblyRibbonManager(Application inventorApp, UserInterfaceManager UIManager, string AddInGUID) : base(inventorApp, UIManager, AddInGUID)
		{
		}

		public override void CreateRibbonPanels()
		{
			ribbon = UIManager.Ribbons["Assembly"];
			ribbonTab = ribbon.RibbonTabs["id_TabAssemble"]; //must use internal names of tabs
			panels.Add(ribbonTab.RibbonPanels.Add("Spiroflow", "assemblyPanel", AddInGUID));
		}
	}
}

