using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using Autodesk.Connectivity.WebServices;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Vault.Currency;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;
using Spiroflow_Vault;
using Folder = Autodesk.Connectivity.WebServices.Folder;

namespace SpiroflowVault
{
	public static class VaultFunctions
	{
		private static Connection GetVaultConnection()
		{
			Connection vaultConnection = Connectivity.Application.VaultBase.ConnectionManager.Instance.Connection;

			if (vaultConnection is null)
			{
				Console.WriteLine("Cannot connect to vault");
				return null;
			}

			return vaultConnection;
		}

		/// <summary>
		/// Gets a list of all .iam and .ipt files in a specific folder (designated by folderID)
		/// </summary>
		/// <param name="folderID"></param>
		/// <returns></returns>
		public static ObservableCollection<VaultFileInfo> GetFilenamesFromFolderId(long folderID)
		{

			ObservableCollection<VaultFileInfo> fileList = new ObservableCollection<VaultFileInfo>();

			var vaultConnection = GetVaultConnection();

			if (vaultConnection is null) return null;

			var docService = vaultConnection.WebServiceManager.DocumentService;

			File[] files = docService.GetLatestFilesByFolderId(folderID, true);

			if (files != null)
			{
				foreach (File file in files)
				{
					if (file.Name.Contains(".dwf")) continue;
					if (file.Name.Contains(".iam") || file.Name.Contains(".ipt"))
					{
						Folder newFolder = docService.GetFolderById(folderID);
						string localFolderPath = newFolder.FullName.Replace(@"$/", @"C:\workspace\");
						string localFilePath = $@"{localFolderPath}\{file.Name}";
						string vaultFilePath = $@"{newFolder.FullName}/{file.Name}";
						vaultFilePath.Replace(@"\", @"/");

						string description = "";//GetDescription(file);
						Image thumbnail = GetThumbnailImage(vaultFilePath);

						fileList.Add(new VaultFileInfo(file.Name, localFolderPath, localFilePath, file.Id, description, thumbnail));
					}
				}
			}

			return fileList;
		}

		private static Image GetThumbnailImage(string vaultFilePath)
		{
			List<string> files = new List<string> { vaultFilePath };
			var vaultConnection = GetVaultConnection();
			try
			{
				var latestFiles = vaultConnection.WebServiceManager.DocumentService.FindLatestFilesByPaths(files.ToArray());
				FileIteration fileIteration = new FileIteration(vaultConnection, latestFiles[0]);

				return GetThumbnailImage(fileIteration);
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				return null;
			}
		}

		private static Image GetThumbnailImage(FileIteration fileIteration, int width = 100, int height = 100)
		{

			var vaultConnection = GetVaultConnection();

			try
			{
				var propDefs = vaultConnection.PropertyManager.GetPropertyDefinitions(EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeSystem);
				var thumbnailPropertyDef = propDefs.SingleOrDefault(x => x.Key == "Thumbnail").Value;
				var propSetting = new PropertyValueSettings();
				var thumbInfo = (ThumbnailInfo) vaultConnection.PropertyManager.GetPropertyValue(fileIteration, thumbnailPropertyDef, propSetting);

				return RenderThumbnailToImage(thumbInfo, height, width);
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				return null;
			}
		}

		private static Image RenderThumbnailToImage(ThumbnailInfo thumbInfo, int height, int width)
		{
			// convert the property value to a byte array
			byte[] thumbnailRaw = thumbInfo.Image as byte[];

			if (null == thumbnailRaw || 0 == thumbnailRaw.Length)
				return null;

			Image retImage = null;

			using (System.IO.MemoryStream memStream = new System.IO.MemoryStream(thumbnailRaw))
			{
				using (System.IO.BinaryReader br = new System.IO.BinaryReader(memStream))
				{
					int CF_METAFILEPICT = 3;
					int CF_ENHMETAFILE = 14;

					int clipboardFormatId = br.ReadInt32(); /*int clipFormat =*/
					bool bytesRepresentMetafile = (clipboardFormatId == CF_METAFILEPICT || clipboardFormatId == CF_ENHMETAFILE);
					try
					{

						if (bytesRepresentMetafile)
						{
							// the bytes represent a clipboard metafile

							// read past header information
							br.ReadInt16();
							br.ReadInt16();
							br.ReadInt16();
							br.ReadInt16();

							System.Drawing.Imaging.Metafile mf = new System.Drawing.Imaging.Metafile(br.BaseStream);
							retImage = mf.GetThumbnailImage(width, height, new Image.GetThumbnailImageAbort(GetThumbnailImageAbort), IntPtr.Zero);
						}
						else
						{
							// the bytes do not represent a metafile, try to convert to an Image
							memStream.Seek(0, System.IO.SeekOrigin.Begin);
							Image im = Image.FromStream(memStream, true, false);

							retImage = im.GetThumbnailImage(width, height, new Image.GetThumbnailImageAbort(GetThumbnailImageAbort), IntPtr.Zero);
						}
					}
					catch
					{
					}
				}
			}

			return retImage;
		}

		private static bool GetThumbnailImageAbort()
		{
			return false;
		}


		private static string GetDescription(File file)
		{
			var vaultConnection = GetVaultConnection();

			var items = vaultConnection.WebServiceManager.ItemService.GetItemsByFileId(file.Id);

			return null;

		}

		/// <summary>
		/// Returns the fileID of a specific file in a specific folder in the Vault. Returns ID of file if found, 0 if not found.
		/// </summary>
		/// <param name="vaultPath">Path to search for the file.</param>
		/// <param name="filename">Filename that you are trying to retrieve the ID for.</param>
		/// <returns>Returns ID of file if found, 0 if not found.</returns>
		public static long GetFileIdByPath(string vaultPath, string filename)
		{
			vaultPath = vaultPath.Replace(@"\", @"/");

			var vaultConnection = GetVaultConnection();
			var docService = vaultConnection.WebServiceManager.DocumentService;
			var searchFolder = docService.GetFolderByPath(vaultPath);

			if (searchFolder is null) return 0;

			File[] files = docService.GetLatestFilesByFolderId(searchFolder.Id, true);

			if (files is null) return 0;

			foreach (var file in files)
			{
				if (file.Name == filename) return file.Id;
			}

			return 0;
		}

		/// <summary>
		/// Gets a list of filenames of all .iam files in the vault in a certain folder.
		/// </summary>
		/// <param name="vaultPath"></param>
		/// <returns></returns>
		public static List<FolderAndFileInfo> GetAssemblyFileNamesFromPath(string vaultPath)
		{
			var fileList = new List<FolderAndFileInfo>();
			vaultPath = vaultPath.Replace(@"\", @"/");

			var vaultConnection = GetVaultConnection();
			var docService = vaultConnection.WebServiceManager.DocumentService;
			var searchFolder = docService.GetFolderByPath(vaultPath);

			if (searchFolder is null) return null;

			Folder[] folders = docService.GetFoldersByParentId(searchFolder.Id, false);

			if (folders is null) return null;

			foreach (var folder in folders)
			{
				if (folder.Name.Contains("OLD")) continue;
				File[] files = docService.GetLatestFilesByFolderId(folder.Id, true);

				if (files is null) continue;

				foreach (var file in files)
				{
					if (file.Name.Contains(".iam"))
					{
						string localFolderPath = folder.FullName.Replace("$/", @"C:\workspace\");
						string localFilePath = $@"{localFolderPath}\{file.Name}";

						fileList.Add(new FolderAndFileInfo(folder.Name, localFolderPath, file.Name, localFilePath, file.Id));
					}
				}
			}

			return fileList;
		}

		public static List<FolderInfo> GetFolderNames(string vaultPath)
		{
			var folderList = new List<FolderInfo>();

			vaultPath = vaultPath.Replace(@"C:\workspace\", @"$\");
			vaultPath = vaultPath.Replace(@"\", @"/");

			var vaultConnection = GetVaultConnection();
			var docService = vaultConnection.WebServiceManager.DocumentService;
			var searchFolder = docService.GetFolderByPath(vaultPath);


			if (searchFolder is null) return null;

			Folder[] folders = docService.GetFoldersByParentId(searchFolder.Id, false);

			if (folders is null) return null;

			foreach (var folder in folders)
			{
				if (folder.Name.Contains("OLD")) continue;
				folderList.Add(new FolderInfo(folder.Name, folder.Id));
			}

			return folderList;
		}

		/// <summary>
		/// This will download a file and all associated children files.
		/// </summary>
		/// <param name="fileId"></param>
		public static void DownloadFileById(long fileId)
		{
			var vaultConnection = GetVaultConnection();
			var settings = new AcquireFilesSettings(vaultConnection);

			settings.DefaultAcquisitionOption = AcquireFilesSettings.AcquisitionOption.Download;
			settings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
			settings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
			settings.OptionsRelationshipGathering.FileRelationshipSettings.VersionGatheringOption = VersionGatheringOption.Latest;

			var docService = vaultConnection.WebServiceManager.DocumentService;
			var fileIteration = new Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.FileIteration(vaultConnection, docService.GetFileById(fileId));

			try
			{
				settings.AddEntityToAcquire(fileIteration);
				vaultConnection.FileManager.AcquireFiles(settings);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
	}
}
