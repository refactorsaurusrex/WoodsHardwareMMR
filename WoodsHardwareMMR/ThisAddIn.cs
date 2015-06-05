using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using NLog;
using NLog.Config;
using NLog.Targets;
using WoodsHardwareMMR.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WoodsHardwareMMR
{
    public partial class ThisAddIn
    {
		public BackupSheet BackupSheet { get; private set; }

		public static ThisAddIn Instance { get; private set; }

		public static string Name { get; private set; }

		public static Logger Logger { get; private set; }

		public static string LogFilePath { get; private set; }

        void ThisAddIn_Startup(object sender, EventArgs e)
        {
	        try
	        {
		        InitializeLogger();

		        if (Application.Version == "15.0")
			        Globals.Ribbons.WoodsRibbon.tabWoods.Label = "WOODS HARDWARE";

		        Instance = this;
		        Logger = LogManager.GetLogger("MmrLog");

		        Name = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyProductAttribute>().Product;
		        BackupSheet = new BackupSheet(Application);

		        Application.SheetActivate += Application_SheetActivate;
		        Application.WorkbookOpen += Application_WorkbookOpen;
		        Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
		        Application.WorkbookActivate += Application_WorkbookActivate;
	        }
	        catch (Exception ex)
	        {
		        Logger.Error(ex.ToString());
				throw;
	        }
        }

	    void InitializeLogger()
	    {
			LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WoodsHardsware", "MmrLog.log");
			var fileTarget = new FileTarget { Layout = "${longdate} ${uppercase:${level}} ${message}", FileName = LogFilePath };

			var config = new LoggingConfiguration();
			config.AddTarget("MmrLog", fileTarget);

			var rule = new LoggingRule("*", LogLevel.Debug, fileTarget);
			config.LoggingRules.Add(rule);

			LogManager.Configuration = config;
	    }

	    void Application_WorkbookActivate(Excel.Workbook workbook)
		{
			try
			{
				var activeSheet = workbook.ActiveSheet as Excel.Worksheet;

				if (activeSheet != null)
					Globals.Ribbons.WoodsRibbon.buttonMmr.Enabled = !activeSheet.Names.Exists(Settings.Default.MmrId);
			}
			catch (Exception ex)
			{
				Logger.Error(ex.ToString());
				throw;
			}
		}

		void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
		{
			try
			{
				if (Application.Workbooks.Count == 1)
					Globals.Ribbons.WoodsRibbon.buttonMmr.Enabled = false;
			}
			catch (Exception ex)
			{
				Logger.Error(ex.ToString());
				throw;
			}
		}

		void Application_WorkbookOpen(Excel.Workbook workbook)
		{
			try
			{
				var activeSheet = workbook.ActiveSheet as Excel.Worksheet;

				if (activeSheet != null)
					Globals.Ribbons.WoodsRibbon.buttonMmr.Enabled = !activeSheet.Names.Exists(Settings.Default.MmrId);
			}
			catch (Exception ex)
			{
				Logger.Error(ex.ToString());
				throw;
			}
		}

		void Application_SheetActivate(object sheet)
		{
			try
			{
				var activeSheet = sheet as Excel.Worksheet;

				if (activeSheet == null)
					return;

				Globals.Ribbons.WoodsRibbon.buttonMmr.Enabled = !activeSheet.Names.Exists(Settings.Default.MmrId);
			}
			catch (Exception ex)
			{
				Logger.Error(ex.ToString());
				throw;
			}
		}

	    void ThisAddIn_Shutdown(object sender, EventArgs e)
	    {
		    try
		    {
			    BackupSheet.Delete();
		    }
		    catch (Exception ex)
		    {
				Logger.Error(ex.ToString());
				throw;
		    }
	    }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
