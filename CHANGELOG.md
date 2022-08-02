# Changelog

All notable changes to this project will be documented in this file.


## Released
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

## Released
## 0.1.0 / 2022-06-30

### Added
- TMetric CSV Parser
- TMetric Data Conversions
- Project grouping and summaries
- Excel Report Writer
- Integrated System.CommandLine Handlers

