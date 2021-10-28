using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpiroflowViewModel.Button_Forms
{
	public class TestClass
	{
		public string firstName { get; set; }
		public string lastName;

		public TestClass(string _firstName, string _lastName)
		{
			firstName = _firstName;
			lastName = _lastName;
		}
	}
}
