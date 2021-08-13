using Inventor;
using System;
using System.Drawing;

namespace SpiroflowAddIn.Ribbons
{
	class DrawingRibbonManager
	{
		//private RibbonTab ribbonTab { get; set; }
		private UserInterfaceManager UIManager { get; set; }
		private string AddInGUID { get; set; }
		private Ribbon drawingRibbon { get; set; }
		private RibbonTab drawingRibbonTab { get; set; }
		private RibbonPanel BOMPanel { get; set; }

		public DrawingRibbonManager(UserInterfaceManager UIManager, string AddInGUID)
		{
			this.UIManager = UIManager;
			this.AddInGUID = AddInGUID;
		}

		public void CreateRibbonPanels(Inventor.Application invApp, ButtonDefinition renumberBOMButton)
		{
			drawingRibbon = UIManager.Ribbons["Drawing"];
			drawingRibbonTab = drawingRibbon.RibbonTabs.Add("Spiroflow", "id_Spiroflow_Drawing", AddInGUID);
			BOMPanel = drawingRibbonTab.RibbonPanels.Add("BOM Functions", "bomFunctions", AddInGUID);
			//create print/pdf/dwg button functions

		}

		public ButtonDefinition CreateButtonDef(Inventor.Application invApp, ButtonDefinition buttonDef, Icon icon)
		{
			buttonDef = invApp.CommandManager.ControlDefinitions.AddButtonDefinition("Renumber BOM", "renumberBOM", CommandTypesEnum.kShapeEditCmdType, AddInGUID, "", "", icon, icon);
			return buttonDef;
		}

		public void AddButton(ButtonDefinition renumberBOMButton)
		{
			BOMPanel.CommandControls.AddButton(renumberBOMButton);
		}
	}
}
