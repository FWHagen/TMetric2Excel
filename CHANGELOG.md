# Changelog

All notable changes to this project will be documented in this file.


## Released

## 0.4.0 / 2022-10-03

### Added
- Supported configuration "DoNotCountProjectHours" to be used in company section
  - Multiple project (by Id - such as RCN-99) may be listed in comma-delimited list
  - Output will show "*" for those projects during runtime
  - Marked projects will appear in first columns, before all others
  - Summation formula on Title column updated to exclude marked projects

### Modified
- Updated finding config file - will use the file found in the first location, not last
- Order of config file location is executable directory, AppData\Remote, Current directory

- - - - 
## 0.3.0 / 2022-09-12

### Added
- Supported configuration "RoundUpToWholeHour" to be used in company section
- Output line showing loaded config file

### Modified
- Bugfix to ConfigParser, now defaults to filename "TMetric2Excel.cfg" in executable directory
- Bugfix to ConfigParser, correctly handling casts of objects from config file
- Generic handling of true/false configuration values to be versatile in capitalization and terms (true, yes, 1, etc.)

- - - - 
## 0.2.0 / 2022-08-02

### Added
- ConfigurationService to handle configurations
- AppConfigParser to parse optional configuration file
- Configurations for shading Weekend, PTO, Holiday, and Sick Day rows

### Modified
- Application will search for files in user directories when not found in current 
    - MyDocuments or Desktop directories
    - Subfolder Timekeeping, TMetric2Excel, or TMetric
    - Command-Line option will still override

- - - - 
## 0.1.0 / 2022-06-30

### Added
- TMetric CSV Parser
- TMetric Data Conversions
- Project grouping and summaries
- Excel Report Writer
- Integrated System.CommandLine Handlers

