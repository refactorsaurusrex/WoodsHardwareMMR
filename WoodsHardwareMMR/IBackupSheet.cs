namespace WoodsHardwareMMR
{
	public interface IBackupSheet 
	{
		void RestoreBackup(Microsoft.Office.Interop.Excel.Worksheet destination);
		void CreateBackup(Microsoft.Office.Interop.Excel.Worksheet source);
	}
}