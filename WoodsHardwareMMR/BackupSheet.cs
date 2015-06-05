using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WoodsHardwareMMR
{
	public class BackupSheet : IBackupSheet
	{
		readonly Excel.Application excelApp;
		Excel.Worksheet backupSheet;

		public BackupSheet(Excel.Application excelApp)
		{
			this.excelApp = excelApp;
		}

		public void RestoreBackup(Excel.Worksheet destination)
		{
			if (backupSheet == null)
				throw new InvalidOperationException("No backup currently exists.");

			backupSheet.Cells.Copy(destination.Cells[1, 1]);
			backupSheet.Cells.Clear();
		}

		public void CreateBackup(Excel.Worksheet source)
		{
			if (backupSheet == null)
				CreateSheet();
			else
				backupSheet.Cells.Clear();

			source.Cells.Copy(backupSheet.Cells[1, 1]);
		}

		public void Delete()
		{
			if (backupSheet != null)
			{
				var screenUpdating = ScreenUpdating.Suppress(excelApp);

				try
				{
					excelApp.DisplayAlerts = false;
					backupSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
					backupSheet.Delete();
				}
				finally
				{
					excelApp.DisplayAlerts = true;
					screenUpdating.Restore();
				}

				backupSheet = null;
				excelApp.WorkbookBeforeClose -= excelApp_WorkbookBeforeClose;
			}
		}

		void CreateSheet()
		{
			excelApp.WorkbookBeforeClose += excelApp_WorkbookBeforeClose;
			var screenUpdating = ScreenUpdating.Suppress(excelApp);

			try
			{
				backupSheet = excelApp.ActiveWorkbook.Worksheets.Add();
				backupSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
			}
			finally
			{
				screenUpdating.Restore();
			}
		}

		void excelApp_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
		{
			try
			{
				Excel.Workbook parentBook = backupSheet.Parent as Excel.Workbook;

				if (parentBook != null && workbook.Name == parentBook.Name)
					Delete();
			}
			catch (Exception ex)
			{
				ThisAddIn.Logger.Error(ex.ToString());
				throw;
			}
		}
	}
}
