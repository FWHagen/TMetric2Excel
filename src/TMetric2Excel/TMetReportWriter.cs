using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    internal class TMetReportWriter : Runtime
    {
        internal void CreateReport(string item, IEnumerable<TMetReportRecord> records, string outputpath, bool showExcel = false)
        {
            Printf(item);
            Printf("-------------------------------------------------");
            Printf("Creating Excel instance");
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            var excel = CreateReportFile(app, item, records.First().Date());
            Printf($"  {records.Count():000} records found for {item}");

            int summaryrow = 35;
            int projectCol = 2;
            var projects = records.Select(tm => tm.ProjectCode).Distinct().OrderBy(tm => tm);
            foreach (var proj in projects)
            {
                var projentries = records.Where(tm => tm.ProjectCode == proj);
                Printf($"  --- {proj,7} - {String.Concat(projentries.First().Project, " ").PadRight(30, '-')}");

                TallyProject(excel.ActiveSheet, projentries.First().Project, projentries, projectCol++);
                summaryrow = WriteSummaries(excel.ActiveSheet, projentries.First().Project, projentries, summaryrow);
            }
            Printf("-------------------------------------------------");
            TotalRows(excel.ActiveSheet, projectCol++, records.First().Date());

            string filename = String.Concat(item.Replace(" ","_"), "-", records.First().Date().ToIsoDate().Substring(0,6), ".xlsx");
            if(!String.IsNullOrWhiteSpace(filename))
                filename = Path.Combine(outputpath, filename);
            FileInfo io = new FileInfo(filename);   
            if(io.Exists)
                io.Delete();
            Log($"Saving {io.FullName}");
            // Save the data
            excel.SaveAs(io.FullName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Close application and release resources
            app.Quit();
            Printf("=================================================\n");
        }

        private _Workbook CreateReportFile(_Application app, string item, DateTime start, bool showExcel = false)
        {
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = showExcel;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = start.ToIsoDate().Substring(0, 6);

            worksheet.Cells[1, 1] = "Date";
            worksheet.Cells.Item[1, 1].Font.Bold = true;

            if (start.Day > 1)
                start = start.AddDays(1 - start.Day);

            int mnth = start.Month;
            int row = start.Day + 1;
            while (start.Month == mnth)
            {
                row = start.Day + 1;
                worksheet.Cells[row, 1] = start;
                start = start.AddDays(1);
            }
            worksheet.UsedRange.EntireColumn.AutoFit();

            return workbook;
        }

        private void TallyProject(_Worksheet excel, string projectName, IEnumerable<TMetReportRecord> projectEntries, int col)
        {
            DateTime start = projectEntries.First().Date();
            if (start.Day > 1)
                start = start.AddDays(1 - start.Day);

            excel.Cells[1, col] = projectName;
            excel.Cells.Item[1, col].Font.Bold = true;

            int mnth = start.Month;
            int row = start.Day + 1;
            while (start.Month == mnth)
            {
                row = start.Day + 1;
                excel.Cells[row, col] = projectEntries.Where(tm => tm.Date() == start).Sum(tm => tm.Duration.ToRealOrZero());

                start = start.AddDays(1);
            }
            excel.Columns.Item[(col)].NumberFormat = "0.0";
            string fcel = (char)(col + 64) + "2";
            string ecel = (char)(col + 64) + row.ToString();
            excel.Cells.Item[row + 1, col] = $"=SUM({fcel}:{ecel})";

        }

        private int WriteSummaries(_Worksheet excel, string projectName, IEnumerable<TMetReportRecord> projentries, int summaryrow)
        {

            excel.Cells[summaryrow, 1] = projectName;
            excel.Cells.Item[summaryrow++, 1].Font.Bold = true;

            var tasks = projentries.Select(tm => tm.TimeEntry).Distinct();
            foreach (var task in tasks)
            {
                Printf($"    {projentries.Where(tm => tm.TimeEntry == task).Sum(tm => tm.Duration.ToRealOrZero()):00} - {task}");
                excel.Cells[summaryrow++, 2] = task;
            }

            return ++summaryrow;
        }

        private void TotalRows(_Worksheet excel, int col, DateTime start)
        {
            if (start.Day > 1)
                start = start.AddDays(1 - start.Day);

            excel.Cells[1, col] = "Total";
            excel.Cells.Item[1, col].Font.Bold = true;
            DrawBorderUnder(excel, 1, 1, 1, col);


            int mnth = start.Month;
            int row = start.Day + 1;
            while (start.Month == mnth)
            {
                row = start.Day + 1;

                string fcel = "B" + row.ToString();
                string ecel = (char)(col + 64-1) + row.ToString();
                excel.Cells.Item[row, col] = $"=SUM({fcel}:{ecel})";

                if (start.DayOfWeek == DayOfWeek.Sunday || start.DayOfWeek == DayOfWeek.Saturday)
                {
                    fcel = "A" + row;
                    ecel = (char)(col + 0 + 64) + row.ToString();  // Change 0 to number columns right of Total desired to shade
                    excel.Range[fcel, ecel].Interior.ColorIndex = 15;
                }
                start = start.AddDays(1);
            }
            DrawBorderUnder(excel, row, 1, row, col);

            excel.Columns.Item[(col)].NumberFormat = "0.0";
            string tfcel = (char)(col + 64) + "2";
            string tecel = (char)(col + 64) + row.ToString();
            excel.Cells.Item[row + 1, col] = $"=SUM({tfcel}:{tecel})";
            excel.Cells.Item[row + 1, col].Font.Bold = true;
        }

        private void DrawBorderUnder(_Worksheet excel, int beginCellRow, int beginCellCol, int endCellRow = 0, int endCellCol = 0, int colorindex = 1)
        {
            var brc = GetCellId(beginCellRow, beginCellCol);
            var erc = brc;
            if(beginCellRow > 0)
                erc = GetCellId(endCellRow, endCellCol);
            var range = excel.Range[$"{brc}:{erc}"];
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = 2d;
            range.Borders[XlBordersIndex.xlEdgeLeft].ColorIndex = colorindex;
        }

        private string GetCellId(int row, int col)
        {
            // this will only work to col 26: "Z" //
            return (char)(col + 64) + row.ToString();
        }
    }
}
