using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Vault.Currency;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;
using Connectivity.Application.VaultBase;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;

namespace SpiroflowVault
{
	public static class VaultFunctions
	{
		private static Connection GetVaultConnection()
		{
			var vaultConnection = ConnectionManager.Instance.Connection;

			if (vaultConnection is null)
			{
				Console.WriteLine("Cannot connect to vault");
				return null;
			}

			return vaultConnection;
		}

		/// <summary>
		///     Gets a list of all .iam and .ipt files in a specific folder (designated by folderID)
		/// </summary>
		/// <param name="folderID"></param>
		/// <returns></returns>
		public static ObservableCollection<VaultFileInfo> GetFilenamesFromFolderId(long folderID)
		{
			var fileList = new ObservableCollection<VaultFileInfo>();

			var vaultConnection = GetVaultConnection();

			if (vaultConnection is null) return null;

			var docService = vaultConnection.WebServiceManager.DocumentService;

			var files = docService.GetLatestFilesByFolderId(folderID, true);

			if (files != null)
				foreach (var file in files)
				{
					if (file.Name.Contains(".dwf")) continue;
					if (file.Name.Contains(".iam") || file.Name.Contains(".ipt"))
					{
						var newFolder = docService.GetFolderById(folderID);
						var localFolderPath = newFolder.FullName.Replace(@"$/", @"C:\workspace\");
						var localFilePath = $@"{localFolderPath}\{file.Name}";
						var vaultFilePath = $@"{newFolder.FullName}/{file.Name}";
						vaultFilePath.Replace(@"\", @"/");

						var description = GetDescription(vaultFilePath);
						var thumbnail = GetThumbnailImage(vaultFilePath);

						fileList.Add(new VaultFileInfo(file.Name, localFolderPath, localFilePath, file.Id, description, thumbnail));
					}
				}

			return fileList;
		}

		private static Image GetThumbnailImage(string vaultFilePath)
		{
			var fileIteration = GetFileIterationFromPath(vaultFilePath);
			try
			{
				return GetThumbnailImage(fileIteration);
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				return null;
			}
		}

		private static FileIteration GetFileIterationFromPath(string vaultFilePath)
		{
			var files = new List<string> { vaultFilePath };
			var vaultConnection = GetVaultConnection();
			try
			{
				var latestFiles = vaultConnection.WebServiceManager.DocumentService.FindLatestFilesByPaths(files.ToArray());
				var fileIteration = new FileIteration(vaultConnection, latestFiles[0]);

				return fileIteration;
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
				var thumbInfo = (ThumbnailInfo)vaultConnection.PropertyManager.GetPropertyValue(fileIteration, thumbnailPropertyDef, propSetting);

				return RenderThumbnailToImage(thumbInfo, height, width);
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				return null;
			}
		}

		/// <summary>
		///     Takes a thumbnail image and stuffs it into a System.Drawing.Image
		/// </summary>
		/// <param name="thumbInfo"></param>
		/// <param name="height"></param>
		/// <param name="width"></param>
		/// <returns></returns>
		private static Image RenderThumbnailToImage(ThumbnailInfo thumbInfo, int height, int width)
		{
			// convert the property value to a byte array
			var thumbnailRaw = thumbInfo.Image;

			if (null == thumbnailRaw || 0 == thumbnailRaw.Length)
				return null;

			Image retImage = null;

			using (var memStream = new MemoryStream(thumbnailRaw))
			{
				using (var br = new BinaryReader(memStream))
				{
					var CF_METAFILEPICT = 3;
					var CF_ENHMETAFILE = 14;

					var clipboardFormatId = br.ReadInt32(); /*int clipFormat =*/
					var bytesRepresentMetafile = clipboardFormatId == CF_METAFILEPICT || clipboardFormatId == CF_ENHMETAFILE;
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

							var mf = new Metafile(br.BaseStream);
							retImage = mf.GetThumbnailImage(width, height, GetThumbnailImageAbort, IntPtr.Zero);
						}
						else
						{
							// the bytes do not represent a metafile, try to convert to an Image
							memStream.Seek(0, SeekOrigin.Begin);
							var im = Image.FromStream(memStream, true, false);

							retImage = im.GetThumbnailImage(width, height, GetThumbnailImageAbort, IntPtr.Zero);
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

		private static string GetDescription(string vaultFilePath)
		{
			var vaultConnection = GetVaultConnection();
			var fileIteration = GetFileIterationFromPath(vaultFilePath);
			try
			{
				var propDefs = vaultConnection.PropertyManager.GetPropertyDefinitions(EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);
				var descriptionPropertyDef = propDefs.SingleOrDefault(x => x.Key == "Description").Value;
				var propSetting = new PropertyValueSettings();
				var descriptionString = (string)vaultConnection.PropertyManager.GetPropertyValue(fileIteration, descriptionPropertyDef, propSetting);
				return descriptionString;
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
				return null;
			}
		}

		/// <summary>
		///     Returns the fileID of a specific file in a specific folder in the Vault. Returns ID of file if found, 0 if not
		///     found.
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

			var files = docService.GetLatestFilesByFolderId(searchFolder.Id, true);

			if (files is null) return 0;

			foreach (var file in files)
				if (file.Name == filename)
					return file.Id;

			return 0;
		}

		/// <summary>
		///     Gets a list of filenames of all .iam files in the vault in a certain folder.
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

			var folders = docService.GetFoldersByParentId(searchFolder.Id, false);

			if (folders is null) return null;

			foreach (var folder in folders)
			{
				if (folder.Name.Contains("OLD")) continue;
				var files = docService.GetLatestFilesByFolderId(folder.Id, true);

				if (files is null) continue;

				foreach (var file in files)
					if (file.Name.Contains(".iam"))
					{
						var localFolderPath = folder.FullName.Replace("$/", @"C:\workspace\");
						var localFilePath = $@"{localFolderPath}\{file.Name}";

						fileList.Add(new FolderAndFileInfo(folder.Name, localFolderPath, file.Name, localFilePath, file.Id));
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

			var folders = docService.GetFoldersByParentId(searchFolder.Id, false);

			if (folders is null) return null;

			foreach (var folder in folders)
			{
				if (folder.Name.Contains("OLD")) continue;
				folderList.Add(new FolderInfo(folder.Name, folder.Id));
			}

			return folderList;
		}

		/// <summary>
		///     This will download a file and all associated children files.
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
			var item = docService.GetFileById(fileId);
			var fileIteration = new FileIteration(vaultConnection, docService.GetFileById(fileId));

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

		public static List<FolderAndFileInfo> FindFilesByFilename(string filename)
		{
			var vaultConnection = GetVaultConnection();
			var docService = vaultConnection.WebServiceManager.DocumentService;

			var propDefs = vaultConnection.PropertyManager.GetPropertyDefinitions(EntityClassIds.Files, null, PropertyDefinitionFilter.IncludeAll);

			var namePropertyDef = propDefs.SingleOrDefault(x => x.Key == "Name").Value;
			var propSetting = new PropertyValueSettings();


			SrchCond[] searchConditions = { new SrchCond() };
			searchConditions[0].PropDefId = namePropertyDef.Id;
			searchConditions[0].PropTyp = PropertySearchType.SingleProperty;
			searchConditions[0].SrchOper = 3; //3 = Is Exactly (from vault SDK)
			searchConditions[0].SrchTxt = filename;

			var rootFolder = docService.GetFolderRoot();
			long[] folderIDs = { rootFolder.Id };
			var status = new SrchStatus();
			var bookmark = string.Empty;

			var filesFound = docService.FindFilesBySearchConditions(searchConditions, null, folderIDs, true, true, ref bookmark, out status);

			var infoList = new List<FolderAndFileInfo>();

			if (filesFound != null && filesFound.Length > 0)
			{
				foreach (var file in filesFound)
				{
					var folder = docService.GetFolderById(file.FolderId);
					var localFolderPath = folder.FullName.Replace(@"$/", @"C:\workspace\");
					var localFilePath = $@"{localFolderPath}\{file.Name}";

					infoList.Add(new FolderAndFileInfo(folder.Name, localFolderPath, file.Name, localFilePath, file.Id));
				}
			}
			else
			{
				MessageBox.Show($"Error getting file {filename}");
			}

			return infoList;
		}

		/// <summary>
		/// Checks if file exists on disk and downloads it if it doesn't or if copy on hard drive is out of date. Should use full filename with path as parameter.
		/// </summary>
		/// <param name="filename"></param>
		/// <returns></returns>
		public static bool CheckFileAndDownloadIfNecessary(string filename)
		{
			if (System.IO.File.Exists(filename))
			{

				return true;
			}

			//try to download from vault
			filename = System.IO.Path.GetFileName(filename);	//strip off directory info
			
			var fileList = FindFilesByFilename(filename);
			if (fileList.Count > 0)
			{
				VaultFunctions.DownloadFileById(fileList[0].Id);
				return true;
			}
			
			MessageBox.Show($"Tried to download {filename} from Vault but failed. Process will now abort.");
			return false;
		}
	}
}