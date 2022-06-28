using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Navigation;
using System.Xml;
using System.Xml.Serialization;
using Autodesk.Connectivity.WebServices;
using Autodesk.DataManagement.Client.Framework.Internal.Automation;
using Inventor;
using SpiroflowAddIn.Button_Forms;
using SpiroflowAddIn.Utilities;
using SpiroflowVault;

namespace SpiroflowAddIn.Buttons
{
	public class ConfigureFSCButton : IButton, INotifyPropertyChanged
	{
		public Inventor.Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		private List<ScrewData> ScrewData { get; set; }
		public Dictionary<string, List<string>> spiralTypes { get; set; }
		public Dictionary<string, List<string>> spiralMaterials { get; set; }

		#region Properties

		private ObservableCollection<VaultFileInfo> feedRestrictors;
		public ObservableCollection<VaultFileInfo> FeedRestrictors
		{
			get => feedRestrictors;
			set
			{
				feedRestrictors = value;
				OnPropertyChanged();
			}
		}

		private ObservableCollection<VaultFileInfo> inletHousings;
		public ObservableCollection<VaultFileInfo> InletHousings
		{
			get => inletHousings;
			set
			{
				inletHousings = value;
				OnPropertyChanged();
			}
		}

		private ObservableCollection<VaultFileInfo> outletHousings;
		public ObservableCollection<VaultFileInfo> OutletHousings
		{
			get => outletHousings;
			set
			{
				outletHousings = value;
				OnPropertyChanged();
			}
		}

		private ObservableCollection<VaultFileInfo> driveConnections;
		public ObservableCollection<VaultFileInfo> DriveConnections
		{
			get => driveConnections;
			set
			{
				driveConnections = value;
				OnPropertyChanged();
			}
		}

		private ObservableCollection<VaultFileInfo> outletChutes;
		public ObservableCollection<VaultFileInfo> OutletChutes
		{
			get => outletChutes;
			set
			{
				outletChutes = value;
				OnPropertyChanged();
			}
		}

		#endregion

		public List<string> ScrewTypes => new List<string>() { "F21 - 214 FLEX SCREW", "F25 - 258 FLEX SCREW", "F31 - 318 FLEX SCREW", "F41 - 412 FLEX SCREW", "F65 - 658 FLEX SCREW", "F83 - 834 FLEX SCREW" };

		public AssemblyDocument assemblyDoc { get; set; }
		private bool HasSetupHappened { get; set; }

		const string XMLDataPath = @"C:\workspace\Designs\START TEMPLATES\FSC\Universal FSC\";

		public event PropertyChangedEventHandler PropertyChanged;

		private ConfigureFSCForm form { get; set; }

		public ConfigureFSCButton()
		{
			DisplayName = $"Configure{System.Environment.NewLine}FSC";
			InternalName = "ConfigureFSCButton";
			PanelID = "assemblyModelPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.test, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			var document = GetInventorApp.GetApp().ActiveDocument;

			if (document == null || document.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			assemblyDoc = (AssemblyDocument) document;
			
			ScrewData = LoadScrewDataFromXML<List<ScrewData>>();

			HasSetupHappened = false;

			form = new ConfigureFSCForm(this);
			//form.DataContext = this;
			form.TypeComboBox.SelectedIndex = 3;
			var helper = new WindowInteropHelper(form);
			helper.Owner = new IntPtr(GetInventorApp.GetApp().MainFrameHWND);
			form.Show();
		}

		private void SetupConfigureForm()
		{
			ReloadComboBoxes(GetCurrentConveyorSize());

			form.LengthTextBox.Text = GetCurrentLength();
			form.OutletAngleTextBox.Text = GetCurrentOutletAngle();
			form.ScrewTypeComboBox.SelectedItem = GetScrewType();
			form.ScrewMaterialComboBox.SelectedItem = GetCurrentScrewMaterial();
			
			form.TubeMaterialComboBox.ItemsSource = new List<string>() {"UHMW", "ST.ST.304 PROSCREW", "CARBON STEEL", "ST.ST.304 OVERSIZED PROSCREW", "UHMW ANTI-STATIC", "UHMW STATIC DISSIPATIVE"};
			form.TubeMaterialComboBox.SelectedItem = GetCurrentTubeMaterial();

			form.CentercoreMaterialComboBox.ItemsSource = new List<string>() {"UHMW, WHITE", "UHMW, BLACK"};
			form.CentercoreMaterialComboBox.SelectedItem = GetCurrentCenterCoreMaterial();
		}

		public void TypeChanged()
		{
			var type = form.TypeComboBox.SelectedItem.ToString();
			if (ScrewTypes.Contains(type)) ReloadComboBoxes(type);
		}

		public string GetCurrentConveyorSize()
		{
			return assemblyDoc.ComponentDefinition.Parameters["size"].Value.ToString();
		}

		public string GetCurrentLength()
		{
			var lengthInCentimeters = assemblyDoc.ComponentDefinition.Parameters["length"].Value;
			return ConvertLengthToInches((lengthInCentimeters/2.54).ToString());
		}

		private string GetCurrentOutletAngle()
		{
			var angleInRadians = assemblyDoc.ComponentDefinition.Parameters["outletAngle"].Value;
			return ((180/Math.PI) * angleInRadians).ToString();				//convert to degrees
		}

		private string GetCurrentTubeMaterial()
		{
			PartComponentDefinition tubeDef = null;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				if (occurrence.DefinitionDocumentType != DocumentTypeEnum.kPartDocumentObject) continue;
				PartDocument partDoc = (PartDocument)occurrence.Definition.Document;

				var designPropSet = partDoc.PropertySets["Design Tracking Properties"];
				string description = designPropSet["Description"].Value.ToString().ToUpper();

				if (description.Contains("CONVEYOR TUBE")) tubeDef = (PartComponentDefinition)occurrence.Definition;

			}

			if (tubeDef != null) return tubeDef.Parameters["type"].Value.ToString();

			return "";
		}

		private string GetCurrentCenterCoreMaterial()
		{
			PartComponentDefinition centerCoreDef = null;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				if (occurrence.DefinitionDocumentType != DocumentTypeEnum.kPartDocumentObject) continue;
				PartDocument partDoc = (PartDocument)occurrence.Definition.Document;

				var designPropSet = partDoc.PropertySets["Design Tracking Properties"];
				string description = designPropSet["Description"].Value.ToString().ToUpper();

				if (description.Contains("CENTER CORE")) centerCoreDef = (PartComponentDefinition)occurrence.Definition;

			}

			if (centerCoreDef != null) return centerCoreDef.Parameters["material"].Value.ToString();

			return "";
		}

		private string GetScrewType()
		{
			PartComponentDefinition screwDef = null;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				if (occurrence.DefinitionDocumentType != DocumentTypeEnum.kPartDocumentObject) continue;
				PartDocument partDoc = (PartDocument) occurrence.Definition.Document;

				var designPropSet = partDoc.PropertySets["Design Tracking Properties"];
				string description = designPropSet["Description"].Value.ToString().ToUpper();

				if (description.Contains("SCREW") && !description.Contains("ADAPTER")) screwDef = (PartComponentDefinition)occurrence.Definition;
			}

			if (screwDef == null) return "";

			switch ((int)screwDef.Parameters["screwTypeID"].Value)
			{
				case 1:
					return "ROUND";
				case 2:
					return "HD ROUND";
				case 3:
					return "FLAT";
				case 4:
					return "PROSCREW OVERSIZED";
				case 5:
					return "PROSCREW REGULAR";
				default:
					return "";
			}
		}

		private string GetCurrentScrewMaterial()
		{
			PartComponentDefinition screwDef = null;

			foreach (ComponentOccurrence occurrence in assemblyDoc.ComponentDefinition.Occurrences)
			{
				if (occurrence.DefinitionDocumentType != DocumentTypeEnum.kPartDocumentObject) continue;
				PartDocument partDoc = (PartDocument)occurrence.Definition.Document;

				var designPropSet = partDoc.PropertySets["Design Tracking Properties"];
				string description = designPropSet["Description"].Value.ToString().ToUpper();

				if (description.Contains("SCREW") && !description.Contains("ADAPTER")) screwDef = (PartComponentDefinition)occurrence.Definition;

			}

			if (screwDef == null) return "";

			return screwDef.Parameters["material"].Value.ToString();

		}

		private void ReloadComboBoxes(string type)
		{
			if (!HasSetupHappened)
			{
				HasSetupHappened = true;
				SetupConfigureForm();
			}

			var sizeCode = type.Substring(6, 3); //gets the 318 from F31 - 318, etc.
			var feedRestrictorVaultPath = $"$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/{type}/00 FEED RESTRICTORS/";
			var inletHousingVaultPath = $"$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/{type}/01 INLET HOUSING ASSY/";
			var outletHousingVaultPath = $"$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/{type}/02 OUTLET HOUSING ASSY/";
			var driveConnectionsVaultPath = $"$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/{type}/03 DRIVE CONNECTIONS/";
			var outletChutesVaultPath = sizeCode == "658" || sizeCode == "834" ? "$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/FSC - COMMON/01 - 658-834 DISCHARGE CHUTES/" : "$/Designs/MECHANICAL/FSC/_FSC Sub-assemblies/FSC - COMMON/00 - 214-412 DISCHARGE CHUTES/";

			form.ScrewTypeComboBox.ItemsSource = GetScrewTypes(sizeCode);
			form.ScrewMaterialComboBox.ItemsSource = GetScrewMaterials(sizeCode);

			FeedRestrictors = GetFilesFromVault(feedRestrictorVaultPath);
			InletHousings = GetFilesFromVault(inletHousingVaultPath);
			OutletHousings = GetFilesFromVault(outletHousingVaultPath);
			DriveConnections = GetFilesFromVault(driveConnectionsVaultPath);
			OutletChutes = GetFilesFromVault(outletChutesVaultPath);
		}

		private ObservableCollection<VaultFileInfo> GetFilesFromVault(string vaultPath)
		{
			List<FolderInfo> folders = VaultFunctions.GetFolderNames(vaultPath);

			ObservableCollection<VaultFileInfo> fileList = new ObservableCollection<VaultFileInfo>();

			if (folders == null || folders.Count == 0) return null;
			//folders.RemoveAll(x => x.files.Count == 0);

			foreach (var folder in folders)
			{
				if (folder.folderName.Contains("OLD")) continue;

				folder.files = VaultFunctions.GetFilenamesFromFolderId((folder.folderID));

				if (folder.files == null || folder.files.Count == 0) continue;

				foreach (var file in folder.files)
				{
					if (file.FileName.Contains(".iam"))
					{
						fileList.Add(file);
					}
				}
			}
			
			return fileList;
		}

		public string ConvertLengthToInches(string input)
		{
			var unitsOfMeasure = assemblyDoc.UnitsOfMeasure;
			return (unitsOfMeasure.GetValueFromExpression(input, Inventor.UnitsTypeEnum.kInchLengthUnits) / 2.54).ToString() + " in";
		}

		private ObservableCollection<string> GetScrewTypes(string sizeCode)
		{
			var screwList = ScrewData.Where(x => x.memberName.Contains(sizeCode))
					.Select(x => x.CoilToCompute)
					.Distinct()
					.ToList();

			return new ObservableCollection<string>(screwList);
		}

		private ObservableCollection<string> GetScrewMaterials(string sizeCode)
		{
			var materialList = ScrewData.Where(x => x.memberName.Contains(sizeCode))
				.Select(x => x.material)
				.Distinct()
				.ToList();

			return new ObservableCollection<string>(materialList);
		}

		private static T LoadScrewDataFromXML<T>()
		{
			var XMLSerializer = new XmlSerializer(typeof(T));
			var filePath = $"{XMLDataPath}screw data.xml";

			var vaultPath = XMLDataPath.Replace(@"C:\workspace\", @"$\");
			vaultPath = vaultPath.Replace(@"\", @"/");

			var dataFileID = VaultFunctions.GetFileIdByPath(vaultPath, "screw data.xml");
			VaultFunctions.DownloadFileById(dataFileID);

			var dataFileInfo = new FileInfo(filePath);
			dataFileInfo.IsReadOnly = false;

			FileStream xmlFileStream = null;
			XmlTextReader xmlReader = null;
			T data;

			try
			{
				xmlFileStream = new FileStream(filePath, FileMode.Open);
				xmlReader = new XmlTextReader(xmlFileStream);				//XmlTextReader.Create(xmlFileStream);
				data = (T) XMLSerializer.Deserialize(xmlReader);
			}
			finally
			{
				xmlReader?.Close();
				xmlFileStream?.Close();
			}

			return data;
		}

		protected void OnPropertyChanged(string name = null)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}
	}
}
