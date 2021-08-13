using Inventor;
using Microsoft.Win32;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SpiroflowAddIn.Ribbons;
using SpiroflowAddIn.Utilities;
using SpiroflowAddIn.Buttons;

namespace Spiroflow_Add_In
{
	/// <summary>
	/// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
	/// that all Inventor AddIns are required to implement. The communication between Inventor and
	/// the AddIn is via the methods on this interface.
	/// </summary>
	[GuidAttribute("f3823a1a-b7e3-416a-a4d6-fa78f7e4ed8c")]
	public class SpiroflowAddin : ApplicationAddInServer
	{

		private Application inventorApp;
		private UserInterfaceManager UIManager;
		private string AddInGUID { get; set; }

		#region Button Definitions
		ButtonDefinition renumberBOMButton;

		#endregion

		public SpiroflowAddin()
		{
		}

		#region ApplicationAddInServer Members

		public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
		{
			// This method is called by Inventor when it loads the addin.
			// The AddInSiteObject provides access to the Inventor Application object.
			// The FirstTime flag indicates if the addin is loaded for the first time.

			// Initialize AddIn members.
			inventorApp = addInSiteObject.Application;
			UIManager = inventorApp.UserInterfaceManager;
			AddInGUID = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value.ToUpper();

			DrawingRibbonManager drawingRibbon = new DrawingRibbonManager(UIManager, AddInGUID);
			drawingRibbon.CreateRibbonPanels(inventorApp, renumberBOMButton);
			//create BOM button functions
			var renumberBOMIcon = CreateImageFromIcon.CreateInventorIcon(SpiroflowAddIn.Properties.Resources.test);
			renumberBOMButton = inventorApp.CommandManager.ControlDefinitions.AddButtonDefinition("Renumber BOM", "renumberBOM", CommandTypesEnum.kShapeEditCmdType, AddInGUID, "", "", renumberBOMIcon, renumberBOMIcon);
			renumberBOMButton.OnExecute += renumberBOMButton_OnExecute;
			renumberBOMButton.Enabled = true;
			drawingRibbon.AddButton(renumberBOMButton);
		}

		public void Deactivate()
		{
			// This method is called by Inventor when the AddIn is unloaded.
			// The AddIn will be unloaded either manually by the user or
			// when the Inventor session is terminated

			// TODO: Add ApplicationAddInServer.Deactivate implementation

			// Release objects.
			inventorApp = null;

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

		public void ExecuteCommand(int commandID)
		{
			// Note:this method is now obsolete, you should use the 
			// ControlDefinition functionality for implementing commands.
		}

		public object Automation
		{
			// This property is provided to allow the AddIn to expose an API 
			// of its own to other programs. Typically, this  would be done by
			// implementing the AddIn's API interface in a class and returning 
			// that class object through this property.

			get
			{
				// TODO: Add ApplicationAddInServer.Automation getter implementation
				return null;
			}
		}

		#endregion

		#region Button Executions
		void renumberBOMButton_OnExecute(NameValueMap context)
		{
			var renumberBOM = new RenumberBOMButton();
			renumberBOM.Execute();
		}
		#endregion

	}
}
