using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Spiroflow_Vault
{
	public class VaultFileInfo
	{
		public string FileName { get; set; }
		public string LocalFolderPath { get; set; }
		public string LocalFilePath { get; set; }
		public long Id { get; set; }
		public string Description { get; set; }

		public ImageSource Picture
		{
			get
			{
				var memStream = new MemoryStream();
				thumbnail.Save(memStream, ImageFormat.Bmp);
				memStream.Seek(0, SeekOrigin.Begin);
				var newBitmap = new BitmapImage();
				newBitmap.BeginInit();
				newBitmap.CacheOption = BitmapCacheOption.OnLoad;
				newBitmap.StreamSource = memStream;
				newBitmap.EndInit();
				return newBitmap;
			}
		}

		public Image thumbnail;

		public VaultFileInfo(string _filename, string _localfolderpath, string _localfilepath, long _Id, string _description, Image _thumbnail)
		{
			FileName = _filename;
			LocalFolderPath = _localfolderpath;
			LocalFilePath = _localfilepath;
			Id = _Id;
			Description = _description;
			thumbnail = _thumbnail;
		}
	}
}
