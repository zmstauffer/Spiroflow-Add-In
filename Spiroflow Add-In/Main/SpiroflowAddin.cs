using Inventor;
using Microsoft.Win32;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Office.CustomUI;
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

		#region Drawing Button Definitions
		ButtonDefinition renumberBOMButton;
		ButtonDefinition createPDFandDWGButton;
		ButtonDefinition createPDFButton;
		ButtonDefinition createDWGButton;
		ButtonDefinition updateDrawingTitleBlock;
		ButtonDefinition createBendNotes;
		#endregion

		#region Assembly Button Definitions
		ButtonDefinition changePartNumbertoFilenameButton;
		ButtonDefinition exportStructuredBOMButton;
		ButtonDefinition replaceSubassemblyButton;
		ButtonDefinition setEngineerAndDateIPropertiesButton;			
		ButtonDefinition findMissingFabricationDrawingsButton;
		ButtonDefinition setWeldmentMaterial;
		ButtonDefinition testButton;                                        //this is to put different random things in to test w/out having to worry about creating a new button every time
																			//ButtonDefinition createNewEquipmentButton;						//this button also shown on zero doc ribbon
		#endregion

		#region Part Button Definitions
		ButtonDefinition setEngineerAndDateIProperties_PartButton;
		#endregion

		#region ZeroDoc Button Definitions
		ButtonDefinition showSettingsWindow;
		//ButtonDefinition createNewEquipmentButtonZeroDoc;
		#endregion

		public SpiroflowAddin()
		{
		}

		#region Main Activation

		public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
		{
			// This method is called by Inventor when it loads the addin.
			// The AddInSiteObject provides access to the Inventor Application object.
			// The FirstTime flag indicates if the addin is loaded for the first time.

			// Initialize AddIn members.
			inventorApp = addInSiteObject.Application;
			UIManager = inventorApp.UserInterfaceManager;
			AddInGUID = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value.ToUpper();

			//add new buttons to the drawing ribbon
			CreateSpiroflowDrawingRibbon();

			//add new buttons to assembly ribbon
			CreateSpiroflowAssemblyRibbon();

			//add new buttons to part ribbon
			CreateSpiroflowPartRibbon();

			//add new buttons to zero doc ribbon
			CreateSpiroflowZeroDocRibbon();
		}

		#endregion

		#region Deactivation
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

		#region Custom Ribbon Creation
		private void CreateSpiroflowDrawingRibbon()
		{
			DrawingRibbonManager drawingRibbon = new DrawingRibbonManager(inventorApp, UIManager, AddInGUID);

			drawingRibbon.CreateRibbonPanels();

			drawingRibbon.AddButton(renumberBOMButton, "SpiroflowAddIn.Buttons.RenumberBOMButton");
			drawingRibbon.AddButton(createPDFandDWGButton, "SpiroflowAddIn.Buttons.CreatePDFandDWGButton");
			drawingRibbon.AddButton(createPDFButton, "SpiroflowAddIn.Buttons.CreatePDFButton");
			drawingRibbon.AddButton(createDWGButton, "SpiroflowAddIn.Buttons.CreateDWGButton");
			drawingRibbon.AddButton(updateDrawingTitleBlock, "SpiroflowAddIn.Buttons.UpdateDrawingTitleBlockButton");
			drawingRibbon.AddButton(createBendNotes, "SpiroflowAddIn.Buttons.CreateBendNotesButton");
		}

		private void CreateSpiroflowAssemblyRibbon()
		{
			AssemblyRibbonManager assemblyRibbon = new AssemblyRibbonManager(inventorApp, UIManager, AddInGUID);

			assemblyRibbon.CreateRibbonPanels();

			assemblyRibbon.AddButton(replaceSubassemblyButton, "SpiroflowAddIn.Buttons.ReplaceSubassemblyButton");
			assemblyRibbon.AddButton(exportStructuredBOMButton, "SpiroflowAddIn.Buttons.ExportStructuredBOMButton");
			assemblyRibbon.AddButton(setEngineerAndDateIPropertiesButton, "SpiroflowAddIn.Buttons.SetEngineerAndDateIPropertiesButton");
			assemblyRibbon.AddButton(findMissingFabricationDrawingsButton, "SpiroflowAddIn.Buttons.FindMissingFabricationDrawingsButton");
			assemblyRibbon.AddButton(setWeldmentMaterial, "SpiroflowAddIn.Buttons.SetWeldmentMaterialButton");
			//assemblyRibbon.AddButton(testButton, "SpiroflowAddIn.Buttons.TestButton");
			//assemblyRibbon.AddButton(changePartNumbertoFilenameButton, "SpiroflowAddIn.Buttons.ChangePartNumberToFilenameButton");
			//assemblyRibbon.AddButton(createNewEquipmentButton, "SpiroflowAddIn.Buttons.CreateNewEquipmentButton");
		}

		private void CreateSpiroflowPartRibbon()
		{
			PartRibbonManager partRibbon = new PartRibbonManager(inventorApp, UIManager, AddInGUID);

			partRibbon.CreateRibbonPanels();

			partRibbon.AddButton(setEngineerAndDateIProperties_PartButton, "SpiroflowAddIn.Buttons.SetEngineerAndDateIProperties_PartButton");
		}

		private void CreateSpiroflowZeroDocRibbon()
		{
			ZeroDocRibbonManager zeroDocRibbon = new ZeroDocRibbonManager(inventorApp, UIManager, AddInGUID);

			zeroDocRibbon.CreateRibbonPanels();

			zeroDocRibbon.AddButton(showSettingsWindow, "SpiroflowAddIn.Buttons.ShowSettingsWindowButton");
			//zeroDocRibbon.AddButton(createNewEquipmentButtonZeroDoc, "SpiroflowAddIn.Buttons.CreateNewEquipmentButtonZeroDoc");
		}
		#endregion
	}
}
