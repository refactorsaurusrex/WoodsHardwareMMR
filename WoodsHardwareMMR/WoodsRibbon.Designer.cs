namespace WoodsHardwareMMR
{
	partial class WoodsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public WoodsRibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tabWoods = this.Factory.CreateRibbonTab();
			this.groupTasks = this.Factory.CreateRibbonGroup();
			this.buttonMmr = this.Factory.CreateRibbonButton();
			this.buttonRestoreLastMmr = this.Factory.CreateRibbonButton();
			this.buttonShowLog = this.Factory.CreateRibbonButton();
			this.tabWoods.SuspendLayout();
			this.groupTasks.SuspendLayout();
			// 
			// tabWoods
			// 
			this.tabWoods.Groups.Add(this.groupTasks);
			this.tabWoods.Label = "Woods Hardware";
			this.tabWoods.Name = "tabWoods";
			// 
			// groupTasks
			// 
			this.groupTasks.Items.Add(this.buttonMmr);
			this.groupTasks.Items.Add(this.buttonRestoreLastMmr);
			this.groupTasks.Items.Add(this.buttonShowLog);
			this.groupTasks.Label = "MMR Tasks";
			this.groupTasks.Name = "groupTasks";
			// 
			// buttonMmr
			// 
			this.buttonMmr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonMmr.Image = global::WoodsHardwareMMR.Properties.Resources.code_add_32_shadow;
			this.buttonMmr.Label = "Format MMR";
			this.buttonMmr.Name = "buttonMmr";
			this.buttonMmr.ScreenTip = "Format MMR";
			this.buttonMmr.ShowImage = true;
			this.buttonMmr.SuperTip = "If you click this button, it means Nick is your favorite son! :)";
			this.buttonMmr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMmr_Click);
			// 
			// buttonRestoreLastMmr
			// 
			this.buttonRestoreLastMmr.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonRestoreLastMmr.Enabled = false;
			this.buttonRestoreLastMmr.Image = global::WoodsHardwareMMR.Properties.Resources.Undo;
			this.buttonRestoreLastMmr.Label = "Undo Last Format";
			this.buttonRestoreLastMmr.Name = "buttonRestoreLastMmr";
			this.buttonRestoreLastMmr.ScreenTip = "Undo Last Format";
			this.buttonRestoreLastMmr.ShowImage = true;
			this.buttonRestoreLastMmr.SuperTip = "Pretty self-explanatory, yes? :D";
			this.buttonRestoreLastMmr.Tag = "";
			this.buttonRestoreLastMmr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRestoreLastMmr_Click);
			// 
			// buttonShowLog
			// 
			this.buttonShowLog.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonShowLog.Image = global::WoodsHardwareMMR.Properties.Resources.information_32_shadow;
			this.buttonShowLog.Label = "Show Logs";
			this.buttonShowLog.Name = "buttonShowLog";
			this.buttonShowLog.ScreenTip = "Show Logs";
			this.buttonShowLog.ShowImage = true;
			this.buttonShowLog.SuperTip = "Click here to open the folder containing log files. This is just in case somethin" +
    "g breaks.";
			this.buttonShowLog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowLog_Click);
			// 
			// WoodsRibbon
			// 
			this.Name = "WoodsRibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tabWoods);
			this.tabWoods.ResumeLayout(false);
			this.tabWoods.PerformLayout();
			this.groupTasks.ResumeLayout(false);
			this.groupTasks.PerformLayout();

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tabWoods;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTasks;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMmr;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRestoreLastMmr;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowLog;
	}

	partial class ThisRibbonCollection
	{
		internal WoodsRibbon WoodsRibbon
		{
			get { return this.GetRibbon<WoodsRibbon>(); }
		}
	}
}
