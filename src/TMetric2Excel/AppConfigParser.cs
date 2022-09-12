namespace TMetric2Excel
{
    internal class AppConfigParser : Runtime
    {
        public List<Configuration> Configs { get; set; } = new List<Configuration>();

        internal List<Configuration> ParseConfigFile(string configfile)
        {
            if (!File.Exists(configfile))
                return null;

            Console.WriteLine($"Loading configuration from {configfile}");
            var currentSection = "Application";
            var alllines = File.ReadAllLines(configfile);
            foreach (var rawLine in alllines)
            {
                var line = rawLine.Trim();
                if (line.IsEmpty())
                    continue;

                if (line.StartsWith("#")) //comment
                    continue;

                if (line.Contains("#"))
                    line = line.Substring(0, line.IndexOf("#"));
                line = line.Trim();
                if (line.IsEmpty())
                    continue;

                if (line.StartsWith("[") && line.EndsWith("]")) //section
                {
                    currentSection = line.Substring(1, line.Length - 2);
                    continue;
                }

                var split = line.Split(new[] { '=' }, 2); //actual config line
                if (split.Length != 2)
                    continue; //empty/invalid line

                var currentKey = split[0].Trim();
                var currentValue = split[1].Trim();

                var definition = new Configuration(currentSection, currentKey, currentValue);

                if(definition != null)
                    Configs.Add(definition);
            }

            return Configs;
        }

        public object GetConfig(string name)
        {
            return GetConfig("Application", name);
        }

        public object GetConfig(string section, string name)
        {
            var cfg = Configs.LastOrDefault(cd => cd.Section == section && cd.Name == name);
            if (cfg == null)
                return null;
            return cfg.Value;
        }

        public void SetConfig(string section, string name, object value)
        {
            var cfg = Configs.LastOrDefault(cd => cd.Section == section && cd.Name == name);
            if (cfg == null)
                Configs.Add(new Configuration(section, name, value));
            else
                cfg.Value = value;
        }

        public void SetAppDefault(string name, object value)
        {
            var cfg = GetConfig(name);
            if (cfg != null)
                return;
            SetConfig("Application", name, value);
        }

    }
}