using Inventor;
using SpiroflowAddIn.Buttons;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SpiroflowAddIn.Ribbons
{
	public class BaseRibbon
	{
		//private RibbonTab ribbonTab { get; set; }
		public Application inventorApp { get; set; }
		public UserInterfaceManager UIManager { get; set; }
		public string AddInGUID { get; set; }
		public Ribbon ribbon { get; set; }
		public RibbonTab ribbonTab { get; set; }
		public List<RibbonPanel> panels { get; set; }

		public BaseRibbon(Application inventorApp, UserInterfaceManager UIManager, string AddInGUID)
		{
			this.inventorApp = inventorApp;
			this.UIManager = UIManager;
			this.AddInGUID = AddInGUID;
			panels = new List<RibbonPanel>();
		}

		public virtual void CreateRibbonPanels() { }

		public void AddButton(ButtonDefinition button, string buttonType)
		{
			IButton buttonObj = null;
			try
			{
				var type = GetButtonType(buttonType);
				buttonObj = (IButton)type;
				buttonObj.invApp = inventorApp;
				buttonObj.buttonDef = button;
				buttonObj.buttonDef = inventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(buttonObj.DisplayName, buttonObj.InternalName, CommandTypesEnum.kShapeEditCmdType, AddInGUID, "", "", buttonObj.smallIcon, buttonObj.icon);
				buttonObj.buttonDef.OnExecute += buttonObj.Execute;
				buttonObj.buttonDef.Enabled = true;

				var panel = panels.First(x => x.InternalName == buttonObj.PanelID);

				if (panel != null)
				{
					panel.CommandControls.AddButton(buttonObj.buttonDef, true);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Failed to load button {buttonType}. \r\n{ex.Message}");
			}
		}

		public object GetButtonType(string buttonType)
		{
			Type type = Type.GetType(buttonType);

			if (type != null) return Activator.CreateInstance(type);

			try                                                                         //this helps to fix errors in finding types from different .dll files etc.
			{
				foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
				{
					type = asm.GetType(buttonType);
					return type == null ? null : Activator.CreateInstance(type);
				}
			}
			catch
			{ }

			return null;
		}
	}
}
