using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WoodsHardwareMMR
{
	public class ScreenUpdating
	{
		readonly Excel.Application excelApp;
		readonly bool turnBackOn;

		public static ScreenUpdating Suppress(Excel.Application excelApp)
		{
			return new ScreenUpdating(excelApp);
		}

		ScreenUpdating(Excel.Application excelApp)
		{
			this.excelApp = excelApp;

			if (excelApp.ScreenUpdating)
			{
				excelApp.ScreenUpdating = false;
				turnBackOn = true;
			}
		}

		public void Restore()
		{
			if (turnBackOn)
				excelApp.ScreenUpdating = true;
		}
	}
}
