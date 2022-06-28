using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpiroflowAddIn.Utilities
{
	public class ScrewData
	{
		public string memberName { get; set; }
		public string partNumber { get; set; }
		public string stockNumber { get; set; }
		public double length { get; set; }
		public int typeID { get; set; }
		public double roundDiameter { get; set; }
		public double centerOffset { get; set; }
		public double profileHeight { get; set; }
		public double profileWidth { get; set; }
		public string size { get; set; }
		public string material { get; set; }
		public string materialString { get; set; }
		public double coilHeight { get; set; }
		public double coilPitch { get; set; }
		public string CoilToCompute { get; set; }
	}
}


