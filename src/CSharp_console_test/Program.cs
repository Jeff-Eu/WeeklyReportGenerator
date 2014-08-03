using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.VisualBasic;
using excel4report;

namespace CSharp_console_test
{
    class Program
    {
        static void Main(string[] args)
        {
            // how to manipulate excel. Tutorial's from
            // http://www.dotblogs.com.tw/feeyaorange/archive/2012/04/24/71751.aspx
            ExcelUtils.xlApp = new Excel.Application();

            string path = AppDomain.CurrentDomain.BaseDirectory;
            string source = "Weekly Report.xlsx";
            string dest = "result_report.xlsx";

            if (File.Exists(path + dest))
                Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(dest,
                    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                    Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);

            Excel.Workbook workbook = ExcelUtils.openWorkBook(path + source);
            Excel.Worksheet src_worksheet = ExcelUtils.getWorkSheet(workbook, 1);
            Excel.Worksheet dest_worksheet = ExcelUtils.getWorkSheet(workbook, 2);

            ReportGenerator rg = new ReportGenerator(src_worksheet, dest_worksheet);

            rg.WriteToGapReport(4); // Send report every Wednesday. Data collected from last Thursday to this Wednesday
            rg.WriteToMainPowerReport(1); // Send report every Friday. Data collected from this Monday to this Friday
            rg.WriteToWeeklyReport(1); // Send report every Monday morning. Data collected from last Monday to last Friday

            ExcelUtils.saveAsWorkBook(workbook, path + dest);
            ExcelUtils.closeWorkBook(workbook);

            //Console.WriteLine("Press any key to exit");
            //System.Console.ReadKey();
        }
    }
}
