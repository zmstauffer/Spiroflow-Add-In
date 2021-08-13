using System.Drawing;

namespace SpiroflowAddIn.Utilities
{
	public static class CreateImageFromIcon
	{
		public static stdole.IPictureDisp CreateInventorIcon(Icon icon)
		{
			return PictureDispConverter.ToIPictureDisp(icon);
		}
	}
}
