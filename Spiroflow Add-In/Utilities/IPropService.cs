using Inventor;
using System;
using System.Collections.Generic;
using System.Windows;

namespace SpiroflowAddIn.Utilities
{
	public static class IPropService
	{

		private static readonly Dictionary<string, string> parameterLocationDictionary = new Dictionary<string, string>()
		{
			{"Designer","Design Tracking Properties"},
			{"Project","Design Tracking Properties"},
			{"Subject","Inventor Summary Information"},
			{"Company","Inventor Document Summary Information"},
			{"Part Number","Design Tracking Properties"},
			{"Description", "Design Tracking Properties"}
		};


		public static T GetParameter<T>(string parameterToGet)
		{
			if (String.IsNullOrEmpty(parameterToGet)) return default(T);

			Document doc = GetInventorApp.GetApp().ActiveDocument;
			object returnValue = default(T);

			try
			{
				if (parameterLocationDictionary.ContainsKey(parameterToGet))
				{
					var propSetName = parameterLocationDictionary[parameterToGet];
					var propSet = doc.PropertySets[propSetName];
					returnValue = propSet[parameterToGet].Value;
				}
				else
				{
					returnValue = doc.PropertySets["Inventor User Defined Properties"][parameterToGet].Value;
				}
			}
			catch
			{
				MessageBox.Show($"Didn't find {parameterToGet} in iProperties or had error retrieving it");
				returnValue = default(T);
			}

			return (T)Convert.ChangeType(returnValue, typeof(T));
		}
	}



}
