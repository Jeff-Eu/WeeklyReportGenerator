using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel4report
{
    class ExcelUtils
    {
        public static Excel.Application xlApp = null;

        public static Excel.Workbook openWorkBook(string path)
        {
            return xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        }

        public static void saveAsWorkBook(Excel.Workbook wb, string path)
        {
            wb.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public static void closeWorkBook(Excel.Workbook wb)
        {
            wb.Close(false, Type.Missing, Type.Missing);
        }

        public static Excel.Workbook getWorkBook(int index)
        {
            if (index > 0 && index <= xlApp.Workbooks.Count)
            {
                return xlApp.Workbooks[index];
            }
            return null;
        }

        public static Excel.Worksheet getWorkSheet(Excel.Workbook wb, int index)
        {
            if (index > 0 && index <= wb.Worksheets.Count)
            {
                return wb.Worksheets[index] as Excel.Worksheet;
            }
            return null;
        }
    }
}
