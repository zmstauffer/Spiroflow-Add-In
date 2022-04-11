using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SpiroflowAddIn.Utilities
{
	public class TreeViewNode : ObservableObject
	{
		private ObservableCollection<TreeViewNode> children;

		// Add all of the properties of a node here. In this example,
		// all we have is a name and whether we are expanded.
		public string Name
		{
			get { return _name; }
			set
			{
				if (_name != value)
				{
					_name = value;
					NotifyPropertyChanged();
				}
			}
		}
		private string _name;

		public bool IsExpanded
		{
			get { return _isExpanded; }
			set
			{
				if (_isExpanded != value)
				{
					_isExpanded = value;
					NotifyPropertyChanged();
				}
			}
		}
		private bool _isExpanded;

		// Children are required to use this in a TreeView
		public IList<TreeViewNode> Children { get { return children; } }

		public TreeViewNode(TreeViewNode parent = null)
		{
			children = new ObservableCollection<TreeViewNode>();
			IsExpanded = false;
		}
	}
}
