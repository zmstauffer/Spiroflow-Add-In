using System;
using System.Globalization;
using System.IO;
using System.Windows.Data;

namespace SpiroflowAddIn.Converters
{
	public class SubPathConverter : IValueConverter
	{
		public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
		{
			string path = (string) value;
			path = new DirectoryInfo(path).Name;
			return path;
		}

		public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
		{
			return "";
		}
	}
}