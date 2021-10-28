using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spiroflow_Vault
{
	public class VaultFileInfo
	{
		public string FileName { get; set; }
		public string LocalFolderPath { get; set; }
		public string LocalFilePath { get; set; }
		public long Id { get; set; }

		public VaultFileInfo(string _filename, string _localfolderpath, string _localfilepath, long _Id)
		{
			FileName = _filename;
			LocalFolderPath = _localfolderpath;
			LocalFilePath = _localfilepath;
			Id = _Id;
		}
	}
}
