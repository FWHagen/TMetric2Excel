using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    internal class ConfigurationService : AppConfigParser
    {
        internal List<Configuration> ParseConfigFile(string configfile = "")
        {
            if (String.IsNullOrWhiteSpace(configfile))
                configfile = "TMetric2Excel.cfg";

            if(!File.Exists(configfile))
            {
                var adpath = Path.Combine(Environment.SpecialFolder.ApplicationData.ToString(), configfile);
                var cdpath = Path.Combine(Environment.CurrentDirectory, configfile);
                var bdpath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, configfile);
                if (File.Exists(adpath))
                    configfile = adpath;
                if (File.Exists(cdpath))
                    configfile = cdpath;
                if (File.Exists(bdpath))
                    configfile = bdpath;
            }

            Configs = base.ParseConfigFile(configfile);

            if(Configs == null)
                Configs = new List<Configuration>();

            SetDefaults();
            return Configs;
        }

        private void SetDefaults()
        {
            SetAppDefault("HolidayColorIndex", 20);
            SetAppDefault("PTOColorIndex", 35);
            SetAppDefault("SickDayColorIndex", 19);
            SetAppDefault("WeekendColorIndex", 15);

            SetAppDefault("ShadeHoliday", true);
            SetAppDefault("ShadePTO", true);
            SetAppDefault("ShadeSickday", true);
            SetAppDefault("ShadeWeekend", true);

            // Not yet implemented //
            SetAppDefault("CountHoliday", false);
            SetAppDefault("CountPTO", false);
            SetAppDefault("CountWeekend", true);
            SetAppDefault("AutoWidenColumnsOnTitle", true);

        }
    }
}
