using System.Runtime.InteropServices;
using Inventor;

namespace SpiroflowAddIn.Utilities
{
	class GetInventorApp
	{
        public static Application GetApp()
        {
            Application inventorApplication;
            try
            {
                inventorApplication = (Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch
            {
                //this is probably a serious error...
                return null;
            }

            return inventorApplication;
        }
    }
}
