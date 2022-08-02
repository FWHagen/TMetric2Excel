using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TMetric2Excel
{
    public class Configuration
    {
        public string Section { get; set; }
        public string Name { get; set; }
        public object Value { get; set; }

        public Configuration(string name, object value)
        {
            Section = "Application";
            Name = name;
            Value = value;  
        }

        public Configuration(string section, string name, object value)
        {
            Section = section;
            Name = name;
            Value = value;
        }
    }
}
