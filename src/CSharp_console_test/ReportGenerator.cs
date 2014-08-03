using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace excel4report
{
    class ReportGenerator
    {
        Excel.Worksheet _src_worksheet;
        Excel.Worksheet _dest_worksheet;

        public ReportGenerator(Excel.Worksheet src_worksheet, Excel.Worksheet dest_worksheet)
        {
            _src_worksheet = src_worksheet;
            _dest_worksheet = dest_worksheet;
        }

        public int CountGroupNumber()
        {
            Excel.Range ra;
            int count = 0;
            for (int i = 2; ; i++)
            {
                ra = ((Excel.Range)_src_worksheet.Cells[i, 1]); // row, column
                string value = ra.Text;

                if (value != "")
                    count++;
                else
                    break;
            }

            return count;
        }

        public int CountLeaveDays()
        {
            Excel.Range ra;
            int count = 0;
            int groupNumber = CountGroupNumber();

            bool isDayLeave = true;
            for (int col = 2; col<=6; col++)
            {
                isDayLeave = true;
                for (int row = 2; row < 2 + groupNumber; row++)
                {
                    ra = ((Excel.Range)_src_worksheet.Cells[row, col]); // row, column
                    string value = ra.Text;

                    if (value != "")
                    {
                        isDayLeave = false;
                        break;
                    }
                }
                if (isDayLeave)
                    count++;
            }
            //// filename
            //ra = ((Excel.Range)src_worksheet.Cells[lineIndex + 1, 1]); // row, column
            //ra.Value2 = words[0];

            return count;
        }

        /// <summary>
        /// Note: the [row,1] in sheet 1 is equal to the [row,1] in sheet 2
        /// </summary>
        /// <param name="project_name">containing the project_name</param>
        /// <returns></returns>
        public int GetProjectRow(string project_name)
        {
            int proj_row = 0;

            Excel.Range ra;

            // get gap_proj_row
            for (int row = 2; ; row++)
            {
                ra = ((Excel.Range)_src_worksheet.Cells[row, 1]); // row, column
                string value = ra.Text;

                if (value != "")
                {
                    if (value.ToLower().Contains(project_name))
                        proj_row = row;
                }
                else
                    break;
            }

            return proj_row;
        }

        public int GetReportCol(string report_name)
        {
            int report_col = 0;
            Excel.Range ra;

            // get gap_proj_row
            for (int col = 2; ; col++)
            {
                ra = ((Excel.Range)_dest_worksheet.Cells[1, col]); // row, column
                string value = ra.Text;

                if (value != "")
                {
                    if (value.ToLower().Contains(report_name))
                        report_col = col;
                }
                else
                    break;
            }

            return report_col;
        }

        /// <summary>
        /// startDay is limited to 1~5.
        /// </summary>
        /// <param name="startDay"></param>
        public void WriteToGapReport(int startDay)
        {
            int gap_col = GetReportCol("gap");
            int gap_row = GetProjectRow("gap");
            Excel.Range ra;

            string report = "";

            int day_col = 0;
            for (int i = 0; i < 5; i++)
            {
                day_col = (startDay + i -1) % 5 + 2;

                ra = ((Excel.Range)_src_worksheet.Cells[gap_row, day_col]);
                string content = ra.Text;

                if(content != "")
                    report += (content + "\n");
            }

            ra = ((Excel.Range)_dest_worksheet.Cells[gap_row, gap_col]); // row, column
            ra.Value2 = report;

                //GetProjectRow("gap");
                //countLeaveDays();
                // countGroupNumber();

            Console.WriteLine("");
        }

        private void writeToEachProject(int startDay, string report_col_name, bool contain_project_others)
        {
            Excel.Range ra;
            int report_col = GetReportCol(report_col_name);
            int day_col = 0;
            int proj_others_row = GetProjectRow("others");
            string report = "";
            int k = contain_project_others ? 1 : 0;

            for (int row = 2; row < proj_others_row + k; row++)
            {
                report = "";

                for (int i = 0; i < 5; i++)
                {
                    day_col = (startDay + i - 1) % 5 + 2;

                    ra = ((Excel.Range)_src_worksheet.Cells[row, day_col]);
                    string content = ra.Text;

                    if (content != "")
                        report += (content + "\n");
                }

                ra = ((Excel.Range)_dest_worksheet.Cells[row, report_col]);
                ra.Value2 = report;
            }
        }

        public string GetProjectName(int row)
        {
            Excel.Range ra;
            ra = ((Excel.Range)_dest_worksheet.Cells[row, 1]);

            return (string)ra.Text;
        }

        public void WriteToMainPowerReport(int startDay)
        {
            writeToEachProject(startDay, "main power", false);

            #region Write topic to each project
            Excel.Range ra;
            int report_col = GetReportCol("main power");
            int proj_others_row = GetProjectRow("others");
            string report = "";

            for (int row = 2; row < proj_others_row; row++)
            {
                report = GetProjectName(row) + ":\n";
                ra = ((Excel.Range)_dest_worksheet.Cells[row, report_col]);

                report += (string)ra.Text;

                ra.Value2 = report;
            }
            #endregion

            // Integrate all projects' reports into the cell of 'others' project
            report = "";
            for (int row = 2; row < proj_others_row; row++)
            {
                ra = ((Excel.Range)_dest_worksheet.Cells[row, report_col]);
                report += (string)ra.Text + "\n\n";
            }
            ra = ((Excel.Range)_dest_worksheet.Cells[GetProjectRow("others"), report_col]);
            ra.Value2 = report;

            // clean projects' reports (except the 'others' row)
            for (int row = 2; row < proj_others_row; row++)
            {
                ra = ((Excel.Range)_dest_worksheet.Cells[row, report_col]);
                ra.Value2 = "";
            }

            // calculate the percentage of each project on each day

            // int leaveDays = CountLeaveDays();
        }

        public void WriteToWeeklyReport(int startDay)
        {
            writeToEachProject(startDay, "weekly report", true);
        }

    }
}
