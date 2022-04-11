using System.Configuration;
using System.Linq;

namespace SpiroflowAddIn.Utilities
{
	static class SettingService
	{
		public static string GetSetting(string settingName)
		{
			return DoesSettingExist(settingName) ? Properties.Settings.Default[settingName].ToString() : "";
		}

		private static bool DoesSettingExist(string settingName)
		{
			return Properties.Settings.Default.Properties.Cast<SettingsProperty>().Any(prop => prop.Name == settingName);
		}

		public static void SetSetting(string settingName, string value)
		{
			if (DoesSettingExist(settingName))
			{
				Properties.Settings.Default[settingName] = value;
				Properties.Settings.Default.Save();
			}
		}
	}
}
