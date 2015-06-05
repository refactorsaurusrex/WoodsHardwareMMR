using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WoodsHardwareMMR
{
	public static class Extensions
	{
		public static bool Exists(this Excel.Names names, string name)
		{
			return names.Cast<Excel.Name>().Any(storedName => storedName.Name.EndsWith(name));
		}

		public static Excel.Name Find(this Excel.Names names, string name)
		{
			return names.Cast<Excel.Name>().SingleOrDefault(storedName => storedName.Name.EndsWith(name));
		}
	}
}
