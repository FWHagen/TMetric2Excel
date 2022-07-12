# TMetric Detailed Monthly Report to Excel Writer

Format Monthly Detailed Reports from TMetric.com to Excel reports separated by client and categorized by project.

This utility will group all TMetric time entries by client and generate totals grouped by project.  The table will be exported to Excel for reporting.

## Getting Started

- Download latest release
- Simply place binary file in empty directory
- Export TMetric Detailed Report to same directory
  - Reports -> Detailed Report
  - Select Calendar -> Last Month
  - Export -> As CSV, downloaded to same directory as TMetric2Excel binary
- Run TMetric2Excel.exe


### Usage:
```
  TMetric2Excel [options]

Options:
  -i, --input-file <input-file>    The input file (will override --months-ago).
  -m, --months-ago <months-ago>    Month to process prior to current month. [default: 1]
  -d, --build-dirs                 Create dated sub-directory for output files. [default: True]
  -o, --output-path <output-path>  File directory path where reports will be created.
  --version                        Show version information
  -?, -h, --help                   Show help and usage information
```

### Prerequisites
- .NET 6.0 runtime 
  - [https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-6.0.6-windows-x64-installer](https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-6.0.6-windows-x64-installer)
- [TMetric](https://app.tmetric.com/) account
- All time entries must have a project assigned
It is assumed that the entries are one calendar month of data



## Authors

* **Frank Hagen - CodeHagen.com**

