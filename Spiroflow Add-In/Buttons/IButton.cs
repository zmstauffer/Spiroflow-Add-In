using Inventor;

namespace SpiroflowAddIn.Buttons
{
	public interface IButton
	{
		Application invApp { get; set; }
		string DisplayName { get; set; }
		string InternalName { get; set; }
		string PanelID { get; set; }
		stdole.IPictureDisp icon { get; set; }
		stdole.IPictureDisp smallIcon { get; set; }
		ButtonDefinition buttonDef { get; set; }

		void Execute(NameValueMap context);
	}
}
