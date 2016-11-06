using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoEmail
{
    public class ExcelInterop
    {

        Excel.Application ExcelApp;
        public List<object[,]> ValueArraysList;

        public ExcelInterop()
        {
            ExcelApp = new Excel.Application();
            ValueArraysList = new List<object[,]>();

        }


        public void OpenSpreadsheets(string thisFileName)
        {
            try
            {
                Excel.Workbook workBook = ExcelApp.Workbooks.Open(thisFileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                GetExcelObject(workBook);

                //
                // Clean up.
                //
                workBook.Close(false, thisFileName, null);
                Marshal.ReleaseComObject(workBook);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading excel: " + ex);
            }

        }



        private void GetExcelObject(Excel.Workbook workBookIn)
        {
            int numSheets = workBookIn.Sheets.Count;

            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)workBookIn.Sheets[sheetNum];
                Excel.Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);
                ValueArraysList.Add(valueArray);
            }


        }



    }
}
