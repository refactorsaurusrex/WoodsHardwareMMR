using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace WoodsHardwareMMR
{
	public partial class WoodsRibbon
	{
		void buttonMmr_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				Excel.Application excelApp = e.Control.Context.Application;
				var formatter = new MmrFormatter(excelApp, ThisAddIn.Instance.BackupSheet);

				if (!formatter.IsValidActiveSheet())
				{
					MessageBox.Show("Sorry, but this does not appear to be a valid MMR sheet.", ThisAddIn.Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				formatter.FormatActiveSheet();
				buttonRestoreLastMmr.Enabled = true;
			}
			catch (Exception ex)
			{
				ThisAddIn.Logger.Error(ex.ToString());
				throw;
			}
		}

		void buttonRestoreLastMmr_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				Excel.Application excelApp = e.Control.Context.Application;
				var formatter = new MmrFormatter(excelApp, ThisAddIn.Instance.BackupSheet);
				formatter.UndoFormat();
				buttonRestoreLastMmr.Enabled = false;
			}
			catch (Exception ex)
			{
				ThisAddIn.Logger.Error(ex.ToString());
				throw;
			}
		}

		void buttonShowLog_Click(object sender, RibbonControlEventArgs e)
		{
			try
			{
				var explorerInfo = new ProcessStartInfo("Explorer.exe", Path.GetDirectoryName(ThisAddIn.LogFilePath));
				Process.Start(explorerInfo);
			}
			catch (Exception ex)
			{
				ThisAddIn.Logger.Error(ex.ToString());
				throw;
			}
		}
	}
}
