# Checkmarx SAST Scan Log Parser
Takes a SAST scan log file or a Checkmarx One Scan ID as an input and opens an Excel document with parsed details from the SAST log file.

When a Scan ID is used additional details will be sourced from Checkmarx One including preset, branch name, origin and other data

Note: When uploading a log file sourced from Checkmarx One with the logFile switch please use the -cxOne switch to ensure it is parsed correctly

Has tabs for General Details, Engine Configuration, Predefined File Exclusions, Phases, Files Processed, Results Summary and General Queries 

Excel created is not saved and must be manually saved if required
