using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using WoodsHardwareMMR.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace WoodsHardwareMMR
{
	public class MmrFormatter
	{
		readonly Excel.Application excelApp;
		readonly IBackupSheet backupSheet;

		readonly List<string> retainedHeaders = new List<string>
		{
			"Day", 
			"# Trans", 
			"Promo Sales", 
			"Net Sales",
			"Gross Profit %"
		};

		public MmrFormatter(Excel.Application excelApp, IBackupSheet backupSheet)
		{
			this.excelApp = excelApp;
			this.backupSheet = backupSheet;
		}

		public void FormatActiveSheet()
		{
			if (!IsValidActiveSheet())
				throw new InvalidOperationException("One or more required headers cannot be found in the active sheet.");

			var screenUpdating = ScreenUpdating.Suppress(excelApp);

			try
			{
				excelApp.ActiveWindow.DisplayWorkbookTabs = true;

				Excel.Worksheet reportSheet = excelApp.ActiveWorkbook.ActiveSheet;
				Excel.Range headerRow = reportSheet.UsedRange.Rows[1];

				backupSheet.CreateBackup(reportSheet);
				int columnCount = (int)excelApp.WorksheetFunction.CountA(headerRow);
				int rowCount = (int)excelApp.WorksheetFunction.CountA(reportSheet.UsedRange.Columns[1]) - 1;

				for (int columnIndex = columnCount; columnIndex >= 1; columnIndex--)
				{
					string headerText = reportSheet.Cells[1, columnIndex].Value;
					if (retainedHeaders.Contains(headerText))
					{
						if (IsAggregateColumn(headerText))
						{
							int newColumnIndex = columnIndex + 1;
							reportSheet.UsedRange.Columns[newColumnIndex].EntireColumn.Insert();
							reportSheet.UsedRange.Cells[1, newColumnIndex].Value = "Rolling " + headerText;

							string formula = string.Format("=SUM(R2C{0}:RC[-1])", columnIndex);
							reportSheet.UsedRange.Cells[2, newColumnIndex].Resize(rowCount).FormulaR1C1 = formula;
						}
					}
					else
					{
						reportSheet.UsedRange.Columns[columnIndex].EntireColumn.Delete();
					}
				}

				headerRow = reportSheet.UsedRange.Rows[1];
				headerRow.Font.Bold = true;
				headerRow.Interior.Color = Color.FromArgb(197, 217, 241);
				headerRow.Borders.Weight = Excel.XlBorderWeight.xlThin;
				headerRow.BorderAround2(Weight: Excel.XlBorderWeight.xlMedium);

				reportSheet.UsedRange.Columns.AutoFit();
				reportSheet.Cells[1, 1].Select();
				reportSheet.Names.Add(Name: Settings.Default.MmrId, RefersTo: Settings.Default.MmrId, Visible: false);


				excelApp.ActiveWindow.SplitRow = 1;
				excelApp.ActiveWindow.FreezePanes = true;
			}
			catch (Exception ex)
			{
				ThisAddIn.Logger.Error(ex.ToString());
				UndoFormat();
				throw;
			}
			finally
			{
				screenUpdating.Restore();
			}
		}

		public void UndoFormat()
		{
			var screenUpdating = ScreenUpdating.Suppress(excelApp);

			try
			{
				Excel.Worksheet undoDestination = excelApp.ActiveWorkbook.ActiveSheet;

				if (undoDestination.Names.Count > 0 && undoDestination.Names.Exists(Settings.Default.MmrId))
					undoDestination.Cells.Clear();
				else
					undoDestination = excelApp.ActiveWorkbook.Worksheets.Add();

				backupSheet.RestoreBackup(undoDestination);

				undoDestination.Cells[1, 1].Select();
				undoDestination.UsedRange.Columns.AutoFit();

				Excel.Name mmrId = undoDestination.Names.Find(Settings.Default.MmrId);
				if (mmrId != null)
					mmrId.Delete();

				excelApp.ActiveWindow.FreezePanes = false;
				excelApp.ActiveWindow.SplitRow = 0;
			}
			finally
			{
				screenUpdating.Restore();
			}
		}

		public bool IsValidActiveSheet()
		{
			Excel.Range headerRow = excelApp.ActiveWorkbook.ActiveSheet.UsedRange.Rows[1];
			Array excelHeaders = headerRow.Value;

			return excelHeaders != null && retainedHeaders.All(h => excelHeaders.Cast<string>().Contains(h));
		}

		bool IsAggregateColumn(string headerText)
		{
			switch (headerText)
			{
				case "Promo Sales":
				case "# Trans":
				case "Net Sales":
					return true;

				default:
					return false;
			}
		}
	}
}
