using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    internal class TMetCsvParser : Runtime
    {
        static Regex csvSplit = new Regex("(?:^|,)(\"(?:[^\"])*\"|[^,]*)", RegexOptions.Compiled);

        public static string[] SplitCSV(string input)
        {

            List<string> list = new List<string>();
            string curr = null;
            foreach (Match match in csvSplit.Matches(input))
            {
                curr = match.Value;
                if (0 == curr.Length)
                {
                    list.Add("");
                }

                list.Add(curr.TrimStart(',').Trim('"'));
            }

            return list.ToArray();
        }
        internal List<TMetReportRecord> ParseFile(string tMetDetailedReportFile)
        {
            FileInfo fileInfo = new FileInfo(tMetDetailedReportFile);
            if (!fileInfo.Exists)
            {
                LogError($"File not found: {tMetDetailedReportFile}");
                return null;
            }

            List<TMetReportRecord> results =new List<TMetReportRecord> ();

            foreach (var item in File.ReadAllLines(fileInfo.FullName))
            {
                var mapped = MapRow(item);
                if(mapped != null)
                    results.Add(mapped);
            }

            return results;
        }

        public TMetReportRecord MapRow(string row)
        {
            var items = SplitCSV(row);

            DateTime day = DateTime.Now;
            if (DateTime.TryParse(items[0], out day) && items.Length > 12)
            {
                var result = new TMetReportRecord()
                {
                    Day = items[0],
                    User = items[1],
                    TimeEntry = items[2],
                    Project = items[3],
                    ProjectCode = items[4],
                    Client = items[5],
                    Tags = items[6],
                    WorkType = items[7],
                    StartTime = items[8],
                    EndTime = items[9], 
                    Duration = items[10],
                    BillableAmount = items[11],
                    BillableCurrency = items[12]
                };

                return result;
            }
            return null;
        }
    }
}
