using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    internal class TMetReportRecord
    {
        public string? Day { get; set; }
        public string? User { get; set; }
        public string? TimeEntry { get; set; }
        public string? Project { get; set; }
        public string? ProjectCode { get; set; }
        public string? Client { get; set; }
        public string? Tags { get; set; }
        public string? WorkType { get; set; }
        public string? StartTime { get; set; }
        public string? EndTime { get; set; }
        public string? Duration { get; set; }
        public string? BillableAmount { get; set; }
        public string? BillableCurrency { get; set; }

    }
}
