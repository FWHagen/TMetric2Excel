using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    internal static class TMetDataConverters
    {
        public static DateTime Date(this TMetReportRecord record)
        {
            DateTime result = DateTime.MinValue;
            if (DateTime.TryParse(record.Day, out result))
                return result;

            return DateTime.MinValue;
        }

        public static DateTime TimeStampStart(this TMetReportRecord record)
        {
            DateTime result = DateTime.MinValue;
            if(DateTime.TryParse(record.Day, out result))
            {
                TimeSpan hr = new TimeSpan(0, 0, 0);
                if (TimeSpan.TryParse(record.StartTime, out hr))
                    result = result + hr;
            }

            return result;
        }

        public static DateTime TimeStampEnd(this TMetReportRecord record)
        {
            DateTime result = DateTime.MinValue;
            if (DateTime.TryParse(record.Day, out result))
            {
                TimeSpan hr = new TimeSpan(0, 0, 0);
                if (TimeSpan.TryParse(record.EndTime, out hr))
                    result = result + hr;
            }

            return result;
        }

    }
}
