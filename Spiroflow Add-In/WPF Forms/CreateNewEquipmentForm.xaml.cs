using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Connectivity.Application.VaultBase;
using VDF = Autodesk.DataManagement.Client.Framework.Vault;
using AWS = Autodesk.Connectivity.WebServices;
using SpiroflowViewModel;
using System.Collections.ObjectModel;

namespace SpiroflowAddIn.WPF_Forms
{
	/// <summary>
	/// Interaction logic for CreateNewEquipmentForm.xaml
	/// </summary>
	public partial class CreateNewEquipmentForm : Window
	{
		private ObservableCollection<TreeViewNode> rootNodes;
		public IEnumerable<TreeViewNode> RootNodes { get => rootNodes; }

		public CreateNewEquipmentForm()
		{
			rootNodes = new ObservableCollection<TreeViewNode>();
			this.DataContext = this;
			InitializeComponent();
			LoadFlexibleScrewConveyorAssemblies();
		}

		private void LoadFlexibleScrewConveyorAssemblies()
		{
			VDF.Currency.Connections.Connection vaultConnection = ConnectionManager.Instance.Connection;

			if (vaultConnection == null)
			{
				MessageBox.Show("No Vault Connection");
				return;
			}

			string vaultPath = $"$/Designs/START TEMPLATES/FSC/";
			vaultPath = vaultPath.Replace(@"\", @"/");

			AWS.Folder searchFolder = vaultConnection.WebServiceManager.DocumentService.GetFolderByPath(vaultPath);
			AWS.Folder[] folders = vaultConnection.WebServiceManager.DocumentService.GetFoldersByParentId(searchFolder.Id, false);          //only want top level here

			if (folders != null && folders.Length > 0)
			{
				TreeViewNode baseNode = new TreeViewNode() { Name = "Flexible Screw Conveyors" };
				baseNode.IsExpanded = true;

				foreach (AWS.Folder folder in folders)
				{
					if (folder.Name.Contains("OLD")) continue;

					var files = vaultConnection.WebServiceManager.DocumentService.GetLatestFilesByFolderId(folder.Id, true);
					if (files != null)
					{
						//create category for these files
						TreeViewNode newCategoryNode = new TreeViewNode() { Name = folder.Name};
						foreach(var file in files)
						{
							if(file.Name.Contains(".iam"))
							{
								TreeViewNode newEquipmentNode = new TreeViewNode() { Name = file.Name };
								newCategoryNode.Children.Add(newEquipmentNode);
							}
						}

						baseNode.Children.Add(newCategoryNode);
					}
				}
				rootNodes.Add(baseNode);
			}
		}
	}
}
