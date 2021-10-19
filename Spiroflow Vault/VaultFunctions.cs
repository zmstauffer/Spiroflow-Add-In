﻿using System;
using System.Collections.Generic;
using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Vault.Currency;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;

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
		public static List<FolderAndFileInfo> GetFilenamesFromId(long folderID)
		{

			List<FolderAndFileInfo> fileList = new List<FolderAndFileInfo>();

			var vaultConnection = GetVaultConnection();

			if (vaultConnection is null) return null;

			var docService = vaultConnection.WebServiceManager.DocumentService;

			File[] files = docService.GetLatestFilesByFolderId(folderID, true);

			if (files is not null)
			{
				foreach (File file in files)
				{
					if (file.Name.Contains(".dwf")) continue;
					if (file.Name.Contains(".iam") || file.Name.Contains(".ipt"))
					{
						Folder newFolder = docService.GetFolderById(file.FolderId);
						string localFolderPath = newFolder.FullName.Replace(@"$/", @"C:\workspace\");
						string localFilePath = $@"{localFolderPath}\{file.Name}";

						fileList.Add(new FolderAndFileInfo(newFolder.Name, localFolderPath, file.Name, localFilePath, file.Id));
					}
				}
			}

			return fileList;
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
