using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Core;

using BrokenVstoDemo.Properties;

namespace BrokenVstoDemo
{
	[ComVisible(true)]
	public class RibbonExtender : IRibbonExtensibility
	{
		#region Constants

		private const string ContactRibbonId = "Microsoft.Outlook.Contact";
		private const string ContactCardContextMenuRibbonId = "Microsoft.Mso.IMLayerUI";
		private const string ExplorerRibbonId = "Microsoft.Outlook.Explorer";

		#endregion

		#region Fields

		private IRibbonUI ribbon;

		#endregion

		#region Ribbon Callbacks

		#region CustomRibbon

		public void CustomRibbonUI_Load(IRibbonUI ribbon)
		{
			this.ribbon = ribbon;
		}

		#endregion

		#region MenuItem1

		public void MenuItem1_Action(IRibbonControl ribbonControl)
		{
			MessageBox.Show("Add-in hasn't failed yet. Open & close a calender item or a few more Outlook items.\r\n" +
			                "After a few tries, the icon will disappear and this menu item will be completely unresponsive.");
		}

		#endregion

		#region MenuItem2

		public void MenuItem2_Action(IRibbonControl ribbonControl)
		{
			MessageBox.Show("Add-in hasn't failed yet. Open & close a calender item or a few more Outlook items.\r\n" +
							"After a few tries, the icon will disappear and this menu item will be completely unresponsive.");
		}

		#endregion

		#endregion

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonId)
		{
			if(ribbonId.Equals(RibbonExtender.ContactCardContextMenuRibbonId, StringComparison.InvariantCultureIgnoreCase))
				return Resources.Ribbon1;
			else
				return String.Empty;
		}

		#endregion
	}
}