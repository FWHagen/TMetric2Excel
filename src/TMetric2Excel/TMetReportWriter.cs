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
        public ConfigurationService ConfigSvc { get; internal set; }

        internal void CreateReport(string item, IEnumerable<TMetReportRecord> records, string outputpath, bool showExcel = false)
        {
            Printf(item);
            Printf("-------------------------------------------------");
            Printf("Creating Excel instance");
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            var excel = CreateReportFile(app, item, records.First().Date());
            Printf($"  {records.Count():000} records found for {item}");

            var companyConfigs = ConfigSvc.Configs.Where(cc => cc.Section == item);
            bool roundInt = ConfigIsTrue(companyConfigs, "RoundUpToWholeHour");
            string[] notTallied = GetConfigValues(companyConfigs, "DoNotCountProjectHours");

            if(roundInt)
                records = RoundUp(records);
            if (notTallied != null && notTallied.Length > 0)
                records = MarkRecordsNotTallied(records, notTallied);

            int summaryrow = 35;
            int projectCol = 2;
            int skiptotal = 0;
            if (notTallied != null && notTallied.Length > 0)
            {
                foreach (var proj in notTallied)
                {
                    var projentries = records.Where(tm => tm.ProjectCode == proj);
                    if (projentries.Any())
                    {
                        skiptotal++;
                        Printf($"  --- {proj,7} - {String.Concat(projentries.First().Project, " * ").PadRight(30, '-')}");

                        TallyProject(excel.ActiveSheet, projentries.First().Project, projentries, projectCol++);
                        summaryrow = WriteSummaries(excel.ActiveSheet, projentries.First().Project, projentries, summaryrow);
                    }
                }
            }
            var projects = records.Select(tm => tm.ProjectCode).Distinct().OrderBy(tm => tm);
            foreach (var proj in projects)
            {
                if (notTallied != null && notTallied.Length > 0 && notTallied.Contains(proj)) 
                    continue;
                var projentries = records.Where(tm => tm.ProjectCode == proj);
                Printf($"  --- {proj,7} - {String.Concat(projentries.First().Project, " ").PadRight(30, '-')}");

                TallyProject(excel.ActiveSheet, projentries.First().Project, projentries, projectCol++);
                summaryrow = WriteSummaries(excel.ActiveSheet, projentries.First().Project, projentries, summaryrow);
            }
            Printf("-------------------------------------------------");
            TotalRows(excel.ActiveSheet, records.First().Client, projectCol++, records.First().Date(), records, skiptotal);

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

        private IEnumerable<TMetReportRecord> MarkRecordsNotTallied(IEnumerable<TMetReportRecord> records, string[] notTallied)
        {
            foreach (var item in notTallied)
            {
                foreach (var record in records.Where(rr => rr.ProjectCode == item))
                {
                    record.NotInTotals = true;
                }
            }
            return records;
        }

        private IEnumerable<TMetReportRecord> RoundUp(IEnumerable<TMetReportRecord> records)
        {
            if (records == null)
                return records;
            foreach (var rec in records)
            {
                if (rec.Duration.IsNotEmpty() && 
                    !rec.Duration.EndsWith(".0") &&
                    !rec.Duration.EndsWith(".00") &&
                    !rec.Duration.EndsWith(".000") &&
                    !rec.Duration.EndsWith(".0000") )
                {
                    double duration = 0.0;
                    double billable = 0.0;
                    double rate = 0.0;
                    if (double.TryParse(rec.Duration, out duration) && duration > 0)
                    {
                        double.TryParse(rec.BillableAmount, out billable);
                        if (duration > 0)
                            rate = billable / duration;
                        duration = Math.Ceiling(duration);
                        rec.Duration = duration.ToString("0.0");
                        rec.BillableAmount = (duration * rate).ToString("0.00");
                    }
                }
            }
            return records;
        }

        private bool ConfigIsTrue(IEnumerable<Configuration> companyConfigs, string key)
        {
            if(companyConfigs != null && companyConfigs.Any(cc => cc.Name == key))
            {
                var config = companyConfigs.LastOrDefault(cc => cc.Name == key);
                if (config.Value.ToString().IsNotEmpty())
                    return config.Value.ToString().ToPessimisticBool();
            }

            return false;
        }
        private string[] GetConfigValues(IEnumerable<Configuration> companyConfigs, string name)
        {
            var config = companyConfigs.FirstOrDefault(cc => cc.Name == name);
            if (config != null)
            {
                return config.Value.ToString().Split(',');
            }
            return new List<string>().ToArray();
        }


        private _Workbook CreateReportFile(_Application app, string item, DateTime start, bool showExcel = false)
        {
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.ActiveSheet;
            // see the excel sheet behind the program  
            app.Visible = showExcel;
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
                excel.Cells[row, col] = projectEntries.Where(tm => tm.Date() == start).Sum(tm => tm.Duration?.ToRealOrZero());

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

        private void TotalRows(_Worksheet excel, string client, int col, DateTime start, IEnumerable<TMetReportRecord> records, int skipped = 0)
        {
            if (start.Day > 1)
                start = start.AddDays(1 - start.Day);

            excel.Cells[1, col] = "Total";
            excel.Cells.Item[1, col].Font.Bold = true;
            DrawBorderUnder(excel, 1, 1, 1, col);

            bool isWeShaded = GetIsShaded(client,"Weekend");
            bool isHolShaded = GetIsShaded(client, "Holiday");
            bool isPtoShaded = GetIsShaded(client, "PTO");
            bool isSdayShaded = GetIsShaded(client, "SickDay");
            int mnth = start.Month;
            int row = start.Day + 1;
            while (start.Month == mnth)
            {
                row = start.Day + 1;

                char startcol = (char)(skipped + 64 + 2);  // 'B' + not tallied //
                string fcel = startcol + row.ToString();
                string ecel = (char)(col + 64-1) + row.ToString();
                excel.Cells.Item[row, col] = $"=SUM({fcel}:{ecel})";

                fcel = "A" + row;
                ecel = (char)(col + 0 + 64) + row.ToString();  // Change 0 to number columns right of Total desired to shade
                
                if (isWeShaded && (start.DayOfWeek == DayOfWeek.Sunday || start.DayOfWeek == DayOfWeek.Saturday))
                    excel.Range[fcel, ecel].Interior.ColorIndex = GetShaderColor(client, "WeekendColorIndex", 15);

                ecel = (char)(col + 2 + 64) + row.ToString();  // Change 0 to number columns right of Total desired to shade

                if (isSdayShaded && records != null &&
                    (records.Any(re => re.Date() == start && (re.TimeEntry == "Sick" ||
                                                                re.TimeEntry == "SickDay" ||
                                                                re.TimeEntry == "Sick Day" ||
                                                                re.TimeEntry == "Flex Day" ||
                                                                re.Tags.Contains("Flex") ||
                                                                re.Tags.Contains("Sick")))))
                {
                    excel.Range[fcel, ecel].Interior.ColorIndex = GetShaderColor(client, "SickDayColorIndex", 20);
                    var item = (records.Last(re => re.Date() == start && (re.TimeEntry == "Sick" ||
                                                                re.TimeEntry == "SickDay" ||
                                                                re.TimeEntry == "Sick Day" ||
                                                                re.TimeEntry == "Flex Day" ||
                                                                re.Tags.Contains("Flex") ||
                                                                re.Tags.Contains("Sick"))));
                    excel.Cells.Item[row, col+2] = item.TimeEntry;
                }

                if (isPtoShaded && records != null &&
                    (records.Any(re => re.Date() == start && (re.TimeEntry == "PTO" ||
                                                                re.TimeEntry == "Day Off" ||
                                                                re.TimeEntry == "Vacation" ||
                                                                re.TimeEntry == "Vacation Day" ||
                                                                re.Tags.Contains("Vacation") ||
                                                                re.Tags.Contains("PTO")))))
                {
                    excel.Range[fcel, ecel].Interior.ColorIndex = GetShaderColor(client, "PTOColorIndex", 20);
                    var item = (records.Last(re => re.Date() == start && (re.TimeEntry == "PTO" ||
                                                                re.TimeEntry == "Day Off" ||
                                                                re.TimeEntry == "Vacation" ||
                                                                re.TimeEntry == "Vacation Day" ||
                                                                re.Tags.Contains("Vacation") ||
                                                                re.Tags.Contains("PTO"))));
                    excel.Cells.Item[row, col + 2] = item.TimeEntry;
                }

                if (isHolShaded && records != null &&
                    (records.Any(re => re.Date() == start && (re.TimeEntry == "Holiday" || re.Tags.Contains("Holiday")))))
                {
                    excel.Range[fcel, ecel].Interior.ColorIndex = GetShaderColor(client, "HolidayColorIndex", 20);
                    var item = (records.Last(re => re.Date() == start && (re.TimeEntry == "Holiday" || re.Tags.Contains("Holiday"))));
                    excel.Cells.Item[row, col + 2] = item.TimeEntry;
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

        private bool GetIsShaded(string client, string linetype = "Weekend")
        {
            var csh = ConfigSvc.GetConfig(client, $"Shade{linetype}");
            if(csh != null)
            {
                return csh.ToString().ToOptimisticBool();
            }

            var ash = ConfigSvc.GetConfig($"Shade{linetype}");
            if (ash != null)
            {
                return ash.ToString().ToOptimisticBool();
            }

            return true;
        }

        private int GetShaderColor(string clientname, string shadername, int defaultvalue)
        {
            var cnamcfg = ConfigSvc.GetConfig(clientname, shadername);
            if (cnamcfg != null)
                return (int)cnamcfg.ToString().ToRealOrZero();
            var cfg = ConfigSvc.GetConfig(shadername);
            if (cfg != null)
                return (int)cfg.ToString().ToRealOrZero();
            return defaultvalue;
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
