namespace SpiroflowVault
{
	public class FolderAndFileInfo
	{
		public string FolderName { get; }
		public string LocalFolderPath { get; }
		public string FileName { get; }
		public string LocalFilePath { get; }
		public long Id { get; }

		public FolderAndFileInfo(string folderName, string localFolderPath, string fileName, string localFilePath, long id)
		{
			FolderName = folderName;
			LocalFolderPath = localFolderPath;
			FileName = fileName;
			LocalFilePath = localFilePath;
			Id = id;
		}

		
	}
}