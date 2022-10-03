using System;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.CommandLine.IO;

namespace TMetric2Excel 
{
    internal class Program : Runtime
    {
        private static ConfigurationService ConfigSvc;

        static async Task<int> Main(string[] args)
        {
            ConfigSvc = new ConfigurationService();

            var fileOption = new Option<FileInfo?>(
                name: "--input-file",
                description: "The input file (will override --months-ago).");
            fileOption.AddAlias("-i");
            var monthsOption = new Option<int>(
                name: "--months-ago",
                description: "Month to process prior to current month.",
                getDefaultValue: () => 1);
            monthsOption.AddAlias("-m");
            var dirOption = new Option<bool>(
                name: "--build-dirs",
                description: "Create dated sub-directory for output files.",
                getDefaultValue: () => true);
            dirOption.AddAlias("-d");
            var outputOption = new Option<string?>(
                name: "--output-path",
                description: "File directory path where reports will be created.");
            outputOption.AddAlias("-o");

            var rootCommand = new RootCommand("Console App to read TMetric Detailed Report CSV file and produce client summary reports in Excel.");
            rootCommand.Add(fileOption);
            rootCommand.Add(monthsOption);
            rootCommand.Add(dirOption);
            rootCommand.Add(outputOption);

            rootCommand.SetHandler(async (file, months, dir, output) =>
            {
                await RunProcess(file!, months, dir, output);
            },
                fileOption, monthsOption, dirOption, outputOption);

            return await rootCommand.InvokeAsync(args);
        }

        private static async Task RunProcess(FileInfo inputfile, int months, bool builddirs, string? outputpath)
        {

            await Task.Run(() =>
            {
                //if (System.Diagnostics.Debugger.IsAttached)
                //{
                //    if (inputfile == null || !inputfile.Exists)
                //        inputfile = new FileInfo(@"..\..\..\..\..\tests\detailed_report_20220501_20220531.csv");

                //    if (String.IsNullOrWhiteSpace(outputpath) && !String.IsNullOrEmpty(inputfile.DirectoryName))
                //        outputpath = inputfile.DirectoryName;
                //}

                var assem = System.Reflection.Assembly.GetExecutingAssembly();
                Printf($"TMetric2Excel v{assem.GetName().Version}");
                Printf($"".PadRight(30, '-'));
                var cfgs = ConfigSvc.ParseConfigFile();
                string tMetDetailedReportFile = FindFileFromArgs(inputfile, months);
                Printf($"".PadRight(60, '-'));

                if (String.IsNullOrWhiteSpace(outputpath))
                {
                    outputpath = AppDomain.CurrentDomain.BaseDirectory;
                    if (System.Diagnostics.Debugger.IsAttached)
                        outputpath = (new FileInfo(tMetDetailedReportFile)).DirectoryName;
                }

                var data = new TMetCsvParser().ParseFile(tMetDetailedReportFile);
                if (data != null)
                {
                    Log($"{data.Count} records retrieved");

                    if (builddirs)
                    {
                        var bdoutputpath = Path.Combine(outputpath, data.First().Date().ToString("yyyyMM"));
                        if (!Directory.Exists(bdoutputpath))
                        {
                            try
                            {
                                Directory.CreateDirectory(bdoutputpath);
                                Log($"Created Directory: {(new FileInfo(bdoutputpath)).FullName}");
                            }
                            catch (Exception ex)
                            {
                                bdoutputpath = outputpath;
                                LogError($"Could not create output directory [{bdoutputpath}]:");
                                LogError(ex.Message);
                            }
                        }
                        outputpath = bdoutputpath;
                    }

                    var clients = data.Select(tm => tm.Client).Distinct().OrderBy(tm => tm);
                    foreach (var item in clients)
                    {
                        if (item != null)
                        {
                            Log($"Processing report for {item}");
                            ProcessClient(item, data.Where(tm => tm.Client == item), outputpath);
                        }
                    }
                }
            });
        }

        private static void ProcessClient(string item, IEnumerable<TMetReportRecord> records, string outputpath = "")
        {
            var tmrw = new TMetReportWriter() { ConfigSvc = ConfigSvc };
            tmrw.CreateReport(item, records, outputpath);
        }

        private static string FindFileFromArgs(FileInfo inputfile, int months)
        {
            if (inputfile != null && inputfile.Exists)
            {
                Log($"- Using input file {inputfile.Name}");
                return inputfile.FullName;
            }

            int monthoffset = 0 - months;
            if (months == 0 && DateTime.Today.Day > 20)
                monthoffset++;

            DateTime start = DateTime.Today;
            start = start.AddDays(1 - start.Day);
            start = start.AddMonths(monthoffset);
            DateTime end = start.AddMonths(1).AddDays(-1);
            Log($"Locating TMetric Detailed Report for {start:MMM} {start:yyyy}:");
            string _tMetDetailedReportFile = $"detailed_report_{start.ToIsoDate()}_{end.ToIsoDate()}.csv";

            if (File.Exists(_tMetDetailedReportFile))
            {
                Log($"- Discovered file at {(new FileInfo(_tMetDetailedReportFile)).FullName}");
                return _tMetDetailedReportFile;
            }

            string[] possiblelocs = new string[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Timekeeping", _tMetDetailedReportFile),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "TMetric2Excel", _tMetDetailedReportFile),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "TMetric", _tMetDetailedReportFile),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Timekeeping", _tMetDetailedReportFile),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "TMetric2Excel", _tMetDetailedReportFile),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "TMetric", _tMetDetailedReportFile),

                Path.Combine(@"..\..\..\..\..\tests\LiveData", _tMetDetailedReportFile)
                };

            foreach (var loca in possiblelocs)
            {
                if (File.Exists(loca))
                {
                    _tMetDetailedReportFile = loca;
                    Log($"- Discovered file at {(new FileInfo(_tMetDetailedReportFile)).FullName}");
                    return _tMetDetailedReportFile;
                }

            }

            // For testing //
            _tMetDetailedReportFile = new FileInfo(@"..\..\..\..\..\tests\detailed_report_20220501_20220531.csv").FullName;
            Log($"- Using TEST file at {(new FileInfo(_tMetDetailedReportFile)).FullName}");

            return _tMetDetailedReportFile;
        }

    }
}