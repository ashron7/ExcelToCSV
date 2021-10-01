using Spire.Xls;
using System;
using System.Text;

namespace ExcelToCSV
{
    class Program
    {
        static Workbook workbook = new Workbook();
        static Workbook workbook2 = new Workbook();
        static void Main(string[] args)
        {

            try
            {
                /*
                 * To check working of error execute failure_scenario() and comment success_scenario()
                 */
                failure_Scenario();
                /*
                 * To check working of error execute failure_scenario() and comment success_scenario()
                 */
                success_Scenario();
            }
            catch (Exception ex)
            {

                Worksheet sheet = workbook.Worksheets[0];

                #region Saving Error Details
                sheet.Rows[1].Columns[0].Value = ex.Message;
                sheet.Rows[1].Columns[1].Value = ex.StackTrace.ToString();
                #endregion

                workbook.SaveToFile("Error.xlsx", ExcelVersion.Version2013);
            }
        }

        private static void failure_Scenario()
        {
            workbook.LoadFromFile(@"E:\Snehal_Kadam_Profile_Backup\Desktop\Current Tasks\Projects\ExcelToCSV\ExcelToCSV\ErrorFileFormat.xlsx");
            workbook2.LoadFromFile(@"E:\Snehal_Kadam_Profile_Backup\Desktop\Current Tasks\Projects\ExcelToCSV\ExcelToCSV\source.xlsx");
            Worksheet sheet = workbook2.Worksheets[0];

            #region Check Row and Column Length
            Console.WriteLine("Number of Columns: " + sheet.LastColumn);
            Console.WriteLine("Number of Rows: " + sheet.LastRow);
            #endregion

            #region Change Column Names
            sheet.Rows[0].Columns[3].Value = "Mfr P/N";
            sheet.Rows[0].Columns[4].Value = "Price";
            sheet.Rows[0].Columns[9].Value = "COO";
            #endregion

            #region Setting Default Values
            for (int i = 1; i < sheet.LastRow; i++)
            {
                sheet.Rows[i].Columns[4].Value = (
                    Convert.ToDouble(sheet.Rows[i].Columns[4].Value) +
                    (Convert.ToDouble(sheet.Rows[i].Columns[4].Value) * 20 / 100)
                    ).ToString();
                sheet.Rows[i].Columns[5].Value = sheet.Rows[i].Columns[5].Value == "" ? "TW" : sheet.Rows[i].Columns[5].Value;
                sheet.Rows[i].Columns[8].Value = sheet.Rows[i].Columns[8].Value == "" ? "EA" : sheet.Rows[i].Columns[8].Value;
            }
            #endregion

            #region Delete Columns
            sheet.DeleteColumn(8);
            #endregion

            sheet.SaveToFile("sample.csv", ",", Encoding.UTF8);
        }
        private static void success_Scenario()
        {
            workbook.LoadFromFile(@"E:\Snehal_Kadam_Profile_Backup\Desktop\Current Tasks\Projects\ExcelToCSV\ExcelToCSV\ErrorFileFormat.xlsx");
            workbook2.LoadFromFile(@"E:\Snehal_Kadam_Profile_Backup\Desktop\Current Tasks\Projects\ExcelToCSV\ExcelToCSV\source.xlsx");
            Worksheet sheet = workbook2.Worksheets[0];

            #region Check Row and Column Length
            Console.WriteLine("Number of Columns: " + sheet.LastColumn);
            Console.WriteLine("Number of Rows: " + sheet.LastRow);
            #endregion

            #region Change Column Names
            sheet.Rows[0].Columns[3].Value = "Mfr P/N";
            sheet.Rows[0].Columns[4].Value = "Price";
            sheet.Rows[0].Columns[5].Value = "COO";
            #endregion

            #region Setting Default Values
            for (int i = 1; i < sheet.LastRow; i++)
            {
                sheet.Rows[i].Columns[4].Value = (
                    Convert.ToDouble(sheet.Rows[i].Columns[4].Value) +
                    (Convert.ToDouble(sheet.Rows[i].Columns[4].Value) * 20 / 100)
                    ).ToString();
                sheet.Rows[i].Columns[5].Value = sheet.Rows[i].Columns[5].Value == "" ? "TW" : sheet.Rows[i].Columns[5].Value;
                sheet.Rows[i].Columns[8].Value = sheet.Rows[i].Columns[8].Value == "" ? "EA" : sheet.Rows[i].Columns[8].Value;
            }
            #endregion

            #region Delete Columns
            sheet.DeleteColumn(8);
            #endregion

            sheet.SaveToFile("sample.csv", ",", Encoding.UTF8);
        }
    }
}
