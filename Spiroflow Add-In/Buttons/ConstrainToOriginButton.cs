using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using SpiroflowAddIn.Buttons;
using SpiroflowAddIn.Utilities;

namespace SpiroflowAddIn.Buttons
{
	class ConstrainToOriginButton : IButton
	{
		public Inventor.Application invApp { get; set; }
		public string DisplayName { get; set; }
		public string InternalName { get; set; }
		public string PanelID { get; set; }
		public stdole.IPictureDisp icon { get; set; }
		public stdole.IPictureDisp smallIcon { get; set; }
		public ButtonDefinition buttonDef { get; set; }

		private AssemblyComponentDefinition assemblyComponentDef { get; set; }

		private WorkPlane xyPlane { get; set; }
		private WorkPlane xzPlane { get; set; }
		private WorkPlane yzPlane { get; set; }

		public ConstrainToOriginButton()
		{
			DisplayName = $"Constrain{System.Environment.NewLine}To Origin";
			InternalName = "ConstrainToOriginButton";
			PanelID = "assemblyModelPanel";
			icon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.constrainToOrigin, 32, 32));
			smallIcon = CreateImageFromIcon.CreateInventorIcon(new System.Drawing.Icon(Properties.Resources.constrainToOrigin, 16, 16));
		}

		public void Execute(NameValueMap context)
		{
			if (invApp.ActiveDocument.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject) return;

			AssemblyDocument assemblyDoc = (AssemblyDocument)GetInventorApp.GetApp().ActiveDocument;

			assemblyComponentDef = assemblyDoc.ComponentDefinition;

			ComponentOccurrence occurrenceToMate;
			if (invApp.ActiveDocument.SelectSet.Count != 1)             //check if we already have something selected, make sure it's only 1 thing
			{
				occurrenceToMate = (ComponentOccurrence)invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select item to constrain to origin");
			}
			else
			{
				occurrenceToMate = (ComponentOccurrence)invApp.ActiveDocument.SelectSet[1];             //index starts at 1...
			}

			if (occurrenceToMate is null)
			{
				MessageBox.Show("Nothing selected.", "ERROR");
				return;
			}
			
			//get main assembly planes
			xyPlane = assemblyComponentDef.WorkPlanes["XY Plane"];
			yzPlane = assemblyComponentDef.WorkPlanes["YZ Plane"];
			xzPlane = assemblyComponentDef.WorkPlanes["XZ Plane"];

			if (occurrenceToMate.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
			{
				MateAssembly(occurrenceToMate);
			}
			else
			{
				MatePart(occurrenceToMate);
			}

		}

		private void MatePart(ComponentOccurrence occurrenceToMate)
		{
			//get mating occurrence planes and proxies
			PartComponentDefinition occurrenceDef = (PartComponentDefinition) occurrenceToMate.Definition;
			
			var occurrenceXYplane = occurrenceDef.WorkPlanes["XY Plane"];
			var occurrenceXZplane = occurrenceDef.WorkPlanes["XZ Plane"];
			var occurrenceYZplane = occurrenceDef.WorkPlanes["YZ Plane"];

			//create proxies
			occurrenceToMate.CreateGeometryProxy(occurrenceXYplane, out object xyPlaneProxy);
			occurrenceToMate.CreateGeometryProxy(occurrenceXZplane, out object xzPlaneProxy);
			occurrenceToMate.CreateGeometryProxy(occurrenceYZplane, out object yzPlaneProxy);

			//create mates
			assemblyComponentDef.Constraints.AddFlushConstraint(xyPlaneProxy, xyPlane, 0);
			assemblyComponentDef.Constraints.AddFlushConstraint(xzPlaneProxy, xzPlane,0);
			assemblyComponentDef.Constraints.AddFlushConstraint(yzPlaneProxy, yzPlane, 0);
		}

		private void MateAssembly(ComponentOccurrence occurrenceToMate)
		{
			//get mating occurrence planes and proxies
			AssemblyComponentDefinition occurrenceDef = (AssemblyComponentDefinition)occurrenceToMate.Definition;

			var occurrenceXYplane = occurrenceDef.WorkPlanes["XY Plane"];
			var occurrenceXZplane = occurrenceDef.WorkPlanes["XZ Plane"];
			var occurrenceYZplane = occurrenceDef.WorkPlanes["YZ Plane"];

			//create proxies
			occurrenceToMate.CreateGeometryProxy(occurrenceXYplane, out object xyPlaneProxy);
			occurrenceToMate.CreateGeometryProxy(occurrenceXZplane, out object xzPlaneProxy);
			occurrenceToMate.CreateGeometryProxy(occurrenceYZplane, out object yzPlaneProxy);

			//create mates
			assemblyComponentDef.Constraints.AddFlushConstraint(xyPlaneProxy, xyPlane, 0);
			assemblyComponentDef.Constraints.AddFlushConstraint(xzPlaneProxy, xzPlane, 0);
			assemblyComponentDef.Constraints.AddFlushConstraint(yzPlaneProxy, yzPlane, 0);
		}
	}
}
