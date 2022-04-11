using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SpiroflowVault
{
	public class FolderInfo
	{
		public string folderName { get; set; }
		public long folderID;
		public ObservableCollection<VaultFileInfo> files { get; set; }

		public FolderInfo(string _folderName, long _folderID)
		{
			folderName = _folderName;
			folderID = _folderID;
			files = new ObservableCollection<VaultFileInfo>();
		}
	}
}