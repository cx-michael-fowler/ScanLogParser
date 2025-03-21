#------------------------------------------------------------------------------------------------------------------------------------------
#region Help

<#
.Synopsis
Opens an Excel file with the Parsed results of a SAST Scan Log 

.Description
Takes a scan log file or a Checkmarx One Scan ID as an input and opens an excel with parsed details from the log file
Has tabs for General Details, Engine Configuration, Predefined File Exclusions, Phases, Files, Results Summary and General Queries
When Scan ID is provided will retrieve additional scan details from Checkmarx One
CxOneAPIModule is required when using Scan ID. The folder must be placed into the same location as the script.

NOTE: Excel created is not saved and must be manually saved if required

Usage
Help
    .\ScanLogParser.ps1 -help [<CommonParameters>]
    
Parse Log File
    .\ScanLogParser.ps1 -logPath <string> [<CommonParameters>]

Parse Log from Checkmarx One Scan ID
    .\ScanLogParser.ps1 -scanId <string> [-silentLogin -apiKey <string] [<CommonParameters>]

.Notes
Version:     2.0
Date:        21/03/2025
Written by:  Michael Fowler
Contact:     michael.fowler@checkmarx.com

Change Log
Version    Detail
-----------------
1.0        Original version
2.0        Added functionality to download log from Checkmarx One using given Scan ID
2.1        Updated Parsing of General Queries and Results Summary
  
.PARAMETER help
Display help

.PARAMETER logPath
The file path for the Scan Log to be processed. Use when providing a downloaded log file

.PARAMETER scanId
A Checkmarx One Scan ID which will be used to retrieve the SAST log

.PARAMETER silentLogin
Log into Checkmarx One using the provided API Key. Is optional and if not used a prompt will appear for the key

.PARAMETER apiKey
The API Key used to log into Checkamrx One. Is mandatory with silentLogin

#>

#endregion
#------------------------------------------------------------------------------------------------------------------------------------------
#region Parameters

[CmdletBinding(DefaultParametersetName='Help')] 
Param (

    [Parameter(ParameterSetName='Help',Mandatory=$false, HelpMessage="Display help")]
    [switch]$help,

    [Parameter(ParameterSetName='File',Mandatory=$true, HelpMessage="Enter Full path for scan log")]
    [string]$logPath,

    [Parameter(ParameterSetName='CxOne',Mandatory=$true, HelpMessage="Enter the Checkmarx One Scan ID")]
    [string]$scanId,

    [Parameter(ParameterSetName='CxOne',Mandatory=$false,HelpMessage="Logon silently using provided API Key")]
    [switch]$silentLogin

)
#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Dynamic Parameters

DynamicParam {
    if ($silentLogin) {
        # Define parameter attributes
        $paramAttributes = New-Object -Type System.Management.Automation.ParameterAttribute
        $paramAttributes.Mandatory = $true
        $paramAttributes.HelpMessage = "The API Key used to login"

        # Create collection of the attributes
        $paramAttributesCollect = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
        $paramAttributesCollect.Add($paramAttributes)

        # Create parameter with name, type, and attributes
        $dynParam = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("apiKey", [string], $paramAttributesCollect)

        # Add parameter to parameter dictionary and return the object
        $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
        $paramDictionary.Add("apiKey", $dynParam)
        return $paramDictionary
    }
}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Begin

Begin {

    if ($scanId) { Import-Module $PSScriptRoot\CxOneAPIModule }
    $apiKey = $PSBoundParameters['apiKey']

    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

    $summary = [System.Collections.Generic.List[ResultsSummary]]::New()
    $general = [System.Collections.Generic.List[GeneralQuery]]::New()
    $files = [System.Collections.Generic.List[File]]::New()
    $predefinedExclusions = [System.Collections.Generic.List[Exclusion]]::New()
    $excludeFiles = [System.Collections.Generic.List[String]]::New()
    $phases = [System.Collections.Generic.List[Phase]]::New()
    $details = [LogDetails]::new()
    $config = @{}
    $conn

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Results Summary Class

    class ResultsSummary {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

        [String]$Query
        [String]$Severity
        [String]$Status
        [Int]$Results
        [Nullable[TimeSpan]]$Duration
        [String]$Cwe

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors
        
        ResultsSummary ([String] $line) {

            $line -match "(.*)\s{2,}Severity:\s(.*)\s{2,}(\D+)Results:\s(.*)\s{2,}Duration\s=\s(\d\d:\d\d:\d\d\.\d\d\d)\s{2,}(.*)\s{2,}CxDescription.*"
            $this.Query = $Matches[1]
            $this.Severity = $Matches[2]
            $this.Status = $Matches[3]
            $this.Results = $Matches[4]
            $this.Duration = [TimeSpan]::ParseExact($Matches[5], "hh\:mm\:ss\.fff", $null)
            $this.Cwe = $Matches[6]
        }
        
        #endregion
        #------------------------------------------------------------------------------------------------------------------------------------------------
    }
    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region General Query Class

    class GeneralQuery {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

        [String]$Query
        [String]$Status
        [Int]$Results
        [Nullable[TimeSpan]]$Duration

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors

        GeneralQuery ([String] $line) {
            $out = $line -split "\s{2,}"
            $this.Query = $out[0].Trim()
            $this.Status = $out[1].Trim()
            $this.Results = $out[2].Trim()
            $this.Duration = [TimeSpan]::ParseExact($out[3].Trim(), "hh\:mm\:ss\.fff", $null)
        }

        #endregion
        #--------------------------------------------------------------------------------------------------------------------------------------------
    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Log Details Class

    class LogDetails {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

            [Nullable[datetime]]$Start
            [Nullable[datetime]]$End
            [Nullable[TimeSpan]]$Runtime
            [String]$Version
            [Int]$ProcessorCount
            [String]$AvailableMemory
            [String]$ProjectName
            [String]$ProjectId
            [String]$ScannedLanguages
            [String]$MultiLanguageMode
            [String]$RelativePath
            [Int]$ExcludeFiles
            [Int]$PredefinedExclusions 
            [Int]$TotalFiles
            [Int]$GoodFiles
            [Int]$PartiallyGoodFiles
            [Int]$BadFiles
            [Int]$ParsedLOC
            [Int]$GoodLOC
            [Int]$BadLOC
            [String]$ScanCoverage
            [String]$ScanCoverageLOC

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors

            LogDetails() { }

        #endregion
        #--------------------------------------------------------------------------------------------------------------------------------------------
    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region File Class

    class File {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

        [Nullable[datetime]]$Start
        [Nullable[datetime]]$End
        [Nullable[TimeSpan]]$Runtime
        [String]$FileName

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors

        File([String] $startLine, [String]$finishLine, [String]$fileName) {      
            $this.FileName = $fileName
            $this.Start = [datetime]::parseexact($startLine.substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null)
            if (-NOT [String]::IsNullOrEmpty($finishLine)) {
                $this.End = [datetime]::parseexact($finishLine.substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null)
                $this.Runtime = New-TimeSpan -Start $this.Start -End $this.End
            }
        }

        #endregion
        #--------------------------------------------------------------------------------------------------------------------------------------------
    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Phase Class

    class Phase {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

        [Nullable[datetime]]$Start
        [Nullable[datetime]]$End
        [Nullable[TimeSpan]]$Runtime
        [String]$PhaseName

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors

        Phase([String] $startLine, [String]$finishLine, [String]$phaseName) {      
            $this.PhaseName = $phaseName
            $this.Start = [datetime]::parseexact($startLine.substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null)
            if (-NOT [String]::IsNullOrEmpty($finishLine)) {
                $this.End = [datetime]::parseexact($finishLine.substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null)
                $this.Runtime = New-TimeSpan -Start $this.Start -End $this.End
            }
        }

        #endregion
        #------------------------------------------------------------------------------------------------------------------------------------------------
    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Exclusion Class

    class Exclusion {
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Variables

        [String]$Reason
        [String]$File

        #endregion    
        #--------------------------------------------------------------------------------------------------------------------------------------------
        #region Constructors

        Exclusion([String]$reason, [String]$file) {
            $this.Reason = $reason
            $this.File = $file
        }

        #endregion
        #--------------------------------------------------------------------------------------------------------------------------------------------

    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Functions

    Function GetLogFile {

        
        # Call Logs API and capture redirect
        $uri = "$($conn.BaseUri)/api/logs/$scanId/sast"
        $response = ApiCall { Invoke-WebRequest $uri -Method GET -Headers $conn.Headers -MaximumRedirection 0 } $conn 
        
        #Set header and URI for redirect
        $tempHeader = $response.headers
        $tempHeader.add("Authorization",$conn.Headers["Authorization"])
        $tempHeader.Remove("Connection") | out-null
        $uri = $response.headers["Location"]
        
        #Get log file and return as array
        $response = ApiCall { Invoke-RestMethod $uri -Method GET -Headers $tempHeader} $conn
        return $response.Split([Environment]::NewLine)
    }

    Function ParseLogFile {
        Param (
            [Array]$lines
        )  

        Write-Verbose "Parsing log file" 
      
        for ($i = 0; $i -lt $lines.Count; $i++) {
   
            #Start time
            if ($i -eq 0) { $details.Start = [datetime]::parseexact($lines[$i].substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null) }

            #Version
            if ($i -eq 1) { $details.Version = $lines[$i].substring(17,7) }

            #Available Memory
            if ($lines[$i] -match "^Used memory: (.*)") { $details.AvailableMemory = $Matches[1] }

            #Processor Count
            if ($lines[$i] -match "^Processor Count: (.*)") { $details.ProcessorCount = $Matches[1] }

            #Current Engine Configuration
            if ($lines[$i] -match "Current Engine Configuration from Application") {
                $i += 2
                while (-NOT [String]::IsNullOrEmpty($lines[$i].Trim())) {
                    $values = $lines[$i] -split "="
                    if ($values -eq 2) { $config[$values[0]] = $null }
                    else { $config[$values[0]] = $values[1] }
                    $i++
                }
            }
        
            #Solution Relative Path
            if ($lines[$i]-match "(Solution relative path is: ')(.*)(')") { $details.RelativePath = $Matches[2] }

            #Excluded files count
            if ($lines[$i] -match "Number of exclude files =([0-9]+)") { 
                $details.ExcludeFiles = $Matches[1]
                if ($details.ExcludeFiles -gt 0) {
                    $i++
                    while (-NOT ([String]::IsNullOrEmpty($lines[++$i].Trim()))) {
                        $excludeFiles.Add($lines[$i].Replace($details.RelativePath,""))
                    }
                }
            }

            if ($lines[$i] -match "Begin Predefined File Exclusions") { 
                while ($lines[++$i] -notmatch "Number of excluded files") {
                    $lines[$i] -match "\[Resolving\] - (.*):(.*)" | Out-Null
                    $predefinedExclusions.Add([Exclusion]::new($Matches[1],$Matches[2]))
                }
                $lines[$i] -match "Number of excluded files: ([0-9]+)" | Out-Null
                $details.PredefinedExclusions = $Matches[1]
            }

            #Scanned Languages
            if ($lines[$i] -match "Languages that will be scanned: (.*)") { $details.ScannedLanguages = $Matches[1] }

            #Multi-Language Mode
            if ($lines[$i] -match "MULTI_LANGUAGE_MODE is set") { $details.MultiLanguageMode = $lines[$i] }

            #Project Name and ID
            if ($lines[$i] -match "(Scan Details: ProjectId=')(.*)(',ProjectName=')(.*)(')") { 
                $details.ProjectId = $Matches[2]
                $details.ProjectName = $Matches[4] 
            }

            #Parsed Files
            if ($lines[$i] -match "Started processing file: (.*)") {
                $j = $i + 1
                $add = $true
                $fileName = $Matches[1].Replace($details.RelativePath,"")
                $fileMatch = $Matches[1].Replace("\", "\\")
                # Loop to find completion of file processing. Exit if end of file
                while (-NOT ($lines[$j] -match $fileMatch)) { 
                    if (++$j -eq ($lines.Count - 1)) { 
                        $files.add([File]::new($lines[$i], $null, $fileName))
                        $add = $false
                        break
                    }
                }
                if ($add) { $files.add([File]::new($lines[$i], $lines[$j], $fileName)) }
            }

            #Processing Phases
            if ($lines[$i] -match "Engine Phase \(Start\): (.*)") {
                $j = $i + 1
                $add = $true
                $phaseName = $Matches[1]
                # Loop to find completion of phase processing. Exit if end of file
                while (-NOT ($lines[$j] -match "Engine Phase \( End \): $phaseName")) { 
                    if (++$j -eq ($lines.Count - 1)) { 
                        $phases.add([Phase]::new($lines[$i], $null, $phaseName)) 
                        $add = $false
                        break
                    }
                }
                if ($add) { $phases.add([Phase]::new($lines[$i], $lines[$j], $phaseName)) }
            }

            #Parsing Summary
            if($lines[$i] -match "^---------------------------$") {
                while (-NOT ([String]::IsNullOrEmpty($lines[++$i].Trim()))) {
                    if ($lines[$i] -match "^Total files(.*)") { $details.TotalFiles = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Good files:(.*)") { $details.GoodFiles = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Partially good files:(.*)") { $details.PartiallyGoodFiles = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Bad files:(.*)") { $details.BadFiles = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Parsed LOC:(.*)") { $details.ParsedLOC = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Good LOC:(.*)") { $details.GoodLOC = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Bad LOC:(.*)") { $details.BadLOC = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Scan coverage:(.*)") { $details.ScanCoverage = $Matches[1].Trim() }
                    if ($lines[$i] -match "^Scan coverage LOC:(.*)") { $details.ScanCoverageLOC = $Matches[1].Trim() }
                }
            }

            #Results summary
            if ($lines[$i] -match "Query - (.*)") { $summary.Add([ResultsSummary]::new($Matches[1])) }
   
            #General Queries
            if ($lines[$i] -match "^(-){27}General") {
                while (-NOT ([String]::IsNullOrEmpty($lines[++$i].Trim()))) { $general.Add([GeneralQuery]::new($lines[$i])) }
            }

            # End Time and runtime
            if ($lines[$i] -match "Exit Main") {
                $details.End = [datetime]::parseexact($lines[$i].substring(0,23), 'dd/MM/yyyy HH:mm:ss,FFF', $null)
                $details.Runtime =  [TimeSpan]::ParseExact($lines[$i].Substring(91,16), "hh\:mm\:ss\.fffffff", $null)
            }
        }
    }

    Function WriteDetailsToExcel {

        Write-Verbose "Creating details worksheet" 
        
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Details"

        if ($scanId) {
            write-host "Retrieving additional scan data from Checkmarx One"
            $scan =  (Get-ScansByIds $conn "All" $scanId)[$scanId]
            $details.ProjectName = $scan.ProjectName

            $uri = "$($conn.BaseUri)/api/sast-metadata/$scanId"
            $metadata = ApiCall { Invoke-RestMethod $uri -Method GET -Headers $conn.Headers  } $conn

            write-host "Additional scan data retrieved"
        }
        
        Write-Verbose "Writing data to worksheet"
        
        WriteGeneralDetailsToExcel $worksheet
        WriteParingSummaryToExcel $worksheet
        if ($scanId) { WriteCxOneDetailsToExcel $worksheet $metadata $scan }
       
        #Auto-fit columns
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
        
        Write-Verbose 'Completed writing data to worksheet "Details"'
    }

    Function WriteGeneralDetailsToExcel {
        Param (
            [Microsoft.Office.Interop.Excel.Worksheet]$worksheet
        )

        #Header
        $worksheet.Range("A1:B1").Merge()
        $worksheet.Range("A1:B1").HorizontalAlignment = -4108 #Align Center
        $worksheet.Range("A1:B1") = "Details"
        $worksheet.Range("A1:B1").Font.Bold=$True

        #Start time
        $worksheet.Range("A2") = "Start Time"
        $worksheet.Range("A2").Font.Bold=$True
        $worksheet.Range("B2") = $details.Start

        #End time
        $worksheet.Range("A3") = "End Time"
        $worksheet.Range("A3").Font.Bold=$True
        $worksheet.Range("B3") = $details.End

        #Total Run time
        $worksheet.Range("A4") = "Total Run Time"
        $worksheet.Range("A4").Font.Bold=$True
        $worksheet.Range("B4") = $details.Runtime.ToString("hh\:mm\:ss\:fff")

        #Version
        $worksheet.Range("A5") = "Version"
        $worksheet.Range("A5").Font.Bold=$True
        $worksheet.Range("B5") = $details.Version

        #Processor Count
        $worksheet.Range("A6") = "Processor Count"
        $worksheet.Range("A6").Font.Bold=$True
        $worksheet.Range("B6") = $details.ProcessorCount

        #Memory
        $worksheet.Range("A7") = "Available Memory"
        $worksheet.Range("A7").Font.Bold=$True
        $worksheet.Range("B7") = $details.AvailableMemory

        #Project Name
        $worksheet.Range("A8") = "Project Name"
        $worksheet.Range("A8").Font.Bold=$True
        $worksheet.Range("B8") = $details.ProjectName

        #Project ID
        $worksheet.Range("A9") = "Project ID"
        $worksheet.Range("A9").Font.Bold=$True
        $worksheet.Range("B9") = $details.ProjectId

        #Scanned Languages
        $worksheet.Range("A10") = "Scanned Languages"
        $worksheet.Range("A10").Font.Bold=$True
        $worksheet.Range("B10") = $details.ScannedLanguages

        #Multi-Language Mode
        $worksheet.Range("A11") = "Multi-Language Mode"
        $worksheet.Range("A11").Font.Bold=$True
        $worksheet.Range("B11") = $details.MultiLanguageMode

        #Predefined Excluded Files Count
        $worksheet.Range("A12") = "Predefined File Exclusions"
        $worksheet.Range("A12").Font.Bold=$True
        $worksheet.Range("B12") = $details.PredefinedExclusions

        #Excluded Files Count
        $worksheet.Range("A13") = "Excluded Files Count"
        $worksheet.Range("A13").Font.Bold=$True
        $worksheet.Range("B13") = $details.ExcludeFiles

        #Formatting
        $worksheet.Range("B1:B13").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A1:B13").Borders.LineStyle = 1
    }

    Function WriteParingSummaryToExcel {
        Param (
            [Microsoft.Office.Interop.Excel.Worksheet]$worksheet
        )

        #Header
        $worksheet.Range("D1:E1").Merge()
        $worksheet.Range("D1:E1") = "Parsing Summary"
        $worksheet.Range("D1:E1").HorizontalAlignment = -4108
        $worksheet.Range("D1:E1").Font.Bold=$True
    
        #Total Files
        $worksheet.Range("D2") = "Total Files"
        $worksheet.Range("D2").Font.Bold=$True
        $worksheet.Range("E2") = $details.TotalFiles.ToString("N0")

        #Partially Good Files
        $worksheet.Range("D3") = "Partially Good Files"
        $worksheet.Range("D3").Font.Bold=$True
        $worksheet.Range("E3") = $details.PartiallyGoodFiles.ToString("N0")

        #Bad Files
        $worksheet.Range("D4") = "Bad Files"
        $worksheet.Range("D4").Font.Bold=$True
        $worksheet.Range("E4") = $details.BadFiles.ToString("N0")

        #Parsed LOC
        $worksheet.Range("D5") = "Parsed LOC"
        $worksheet.Range("D5").Font.Bold=$True
        $worksheet.Range("E5") = $details.ParsedLOC.ToString("N0")

        #Good LOC
        $worksheet.Range("D6") = "Good LOC"
        $worksheet.Range("D6").Font.Bold=$True
        $worksheet.Range("E6") = $details.GoodLOC.ToString("N0")

        #Bad LOC
        $worksheet.Range("D7") = "Bad LOC"
        $worksheet.Range("D7").Font.Bold=$True
        $worksheet.Range("E7") = $details.BadLOC.ToString("N0")
    
        #Scan Coverage
        $worksheet.Range("D8") = "Scan Coverage"
        $worksheet.Range("D8").Font.Bold=$True
        $worksheet.Range("E8") = $details.ScanCoverage

        #Scan Coverage LOC
        $worksheet.Range("D9") = "Scan Coverage LOC"
        $worksheet.Range("D9").Font.Bold=$True
        $worksheet.Range("E9") = $details.ScanCoverageLOC

        #Excluded Files
        if ($details.ExcludeFiles -gt 0) {
            $worksheet.Range("G1") = "Excluded Files"
            $worksheet.Range("G1").Font.Bold=$True
            $i = 2
            foreach ($ef in $excludeFiles) { 
                $worksheet.Range("G$i") = $ef
                $i++
            }
            $i--
            $worksheet.Range("G1:G$i").HorizontalAlignment = -4131 #Align Left
            $worksheet.Range("G1:G$i").Borders.LineStyle = 1
        }

        #Formatting
        $worksheet.Range("D1:E9").Borders.LineStyle = 1
    }

    Function WriteCxOneDetailsToExcel {
        Param (
            [Microsoft.Office.Interop.Excel.Worksheet]$worksheet,
            [PSCustomObject]$metadata,
            [Object]$scan
        )

        #Header
        $worksheet.Range("A15:B15").Merge()
        $worksheet.Range("A15:B15").HorizontalAlignment = -4108 #Align Center
        $worksheet.Range("A15:B15") = "Checkmarx One Scan Data"
        $worksheet.Range("A15:B15").Font.Bold=$True

        #Status
        $worksheet.Range("A16") = "Scan Status"
        $worksheet.Range("A16").Font.Bold=$True
        $worksheet.Range("B16") = $scan.Status
        
        #Branch
        $worksheet.Range("A17") = "Branch"
        $worksheet.Range("A17").Font.Bold=$True
        $worksheet.Range("B17") = $scan.Branch
        
        #Preset
        $worksheet.Range("A18") = "Preset"
        $worksheet.Range("A18").Font.Bold=$True
        $worksheet.Range("B18") = $metadata.queryPreset

        #Is Incremental
        $worksheet.Range("A19") = "Incremental Scan"
        $worksheet.Range("A19").Font.Bold=$True
        $worksheet.Range("B19") = $metadata.isIncremental

        #Incremental Cancelled
        $worksheet.Range("A20") = "Incremental Scan Cancelled"
        $worksheet.Range("A20").Font.Bold=$True
        $worksheet.Range("B20") = $metadata.isIncrementalCanceled

        #Incremental Cancelled Reason
        $worksheet.Range("A21") = "Incremental Cancelled Reason"
        $worksheet.Range("A21").Font.Bold=$True
        $worksheet.Range("B21") = $metadata.incrementalCancelReason

        #Initiator
        $worksheet.Range("A22") = "Initiator"
        $worksheet.Range("A22").Font.Bold=$True
        $worksheet.Range("B22") = $scan.Initiator
        
        #Source Type
        $worksheet.Range("A23") = "Source Type"
        $worksheet.Range("A23").Font.Bold=$True
        $worksheet.Range("B23") = $scan.SourceType

        #Source Origin
        $worksheet.Range("A24") = "Source Origin"
        $worksheet.Range("A24").Font.Bold=$True
        $worksheet.Range("B24") = $scan.SourceOrigin

        #Formatting
        $worksheet.Range("B16:B24").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A15:B24").Borders.LineStyle = 1
    }

    Function WritePhasesToExcel {

        Write-Verbose "Creating phases worksheet" 
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "Phases"
    
        Write-Verbose "Writing data to worksheet"

        #Headers
        $worksheet.Range("A1") = "Phase"
        $worksheet.Range("B1") = "Start Time"
        $worksheet.Range("C1") = "End Time"
        $worksheet.Range("D1") = "Run Time"
        $worksheet.Range("A1:D1").Font.Bold=$True

        $i = 2
    
        #Writing Phases List to worksheet
        foreach ($phase in $phases) {
            $worksheet.Range("A$i") = $phase.PhaseName
            $worksheet.Range("B$i") = $phase.Start
            $worksheet.Range("C$i") = $phase.End
            try { $worksheet.Range("D$i") = $phase.Runtime.ToString("hh\:mm\:ss\:fff") }
            catch { $worksheet.Range("D$i") = "" }
            $i++
        }

        #Formatting
        $i--
        $worksheet.Range("A1:D$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A1:D$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "Phases"'
    }

    Function WriteFilesToExcel {

        Write-Verbose "Creating Files worksheet" 
    
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "Files Processed"
    
        Write-Verbose "Writing data to worksheet"

        #Headers
        $worksheet.Range("A1") = "File"
        $worksheet.Range("B1") = "Start Time"
        $worksheet.Range("C1") = "End Time"
        $worksheet.Range("D1") = "Run Time"
        $worksheet.Range("A1:D1").Font.Bold=$True

        $i = 2

        #Writing Files List to worksheet
        foreach ($file in $files) {
            $worksheet.Range("A$i") = $file.FileName
            $worksheet.Range("B$i") = $file.Start
            $worksheet.Range("C$i") = $file.End
            try { $worksheet.Range("D$i") = $file.Runtime.ToString("hh\:mm\:ss\:fff") }
            catch { $worksheet.Range("D$i") = "" }
            $i++
        }
    
        #Formatting
        $i--
        $worksheet.Range("A1:D$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A1:D$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "Files Processed"'
    }

    Function WriteSummaryToExcel {

        Write-Verbose "Creating Results Summary worksheet" 
    
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "Results Summary"
    
        Write-Verbose "Writing data to worksheet"

        #Headers
        $worksheet.Range("A1") = "Query"
        $worksheet.Range("B1") = "Severity"
        $worksheet.Range("C1") = "Status"
        $worksheet.Range("D1") = "Results"
        $worksheet.Range("E1") = "Duration"
        $worksheet.Range("F1") = "CWE"

        $worksheet.Range("A1:F1").Font.Bold=$True

        $i = 2
    
        #Writing summary List to worksheet
        foreach ($summ in $summary) {
            $worksheet.Range("A$i") = $summ.Query
            $worksheet.Range("B$i") = $summ.Severity
            $worksheet.Range("C$i") = $summ.Status
            $worksheet.Range("D$i") = $summ.Results.ToString("N0")
            $worksheet.Range("E$i") = $summ.Duration.ToString("hh\:mm\:ss\:fff")
            $worksheet.Range("F$i") = $summ.Cwe
            $i++
        }

        #Formatting
        $i--
        $worksheet.Range("A1:F$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("D1:D$i").HorizontalAlignment = -4108 #Align Center
        $worksheet.Range("A1:F$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:F$i"))
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "Results Summary"'
    }

    Function WriteGeneralToExcel {
    
        Write-Verbose "Creating Results Summary worksheet" 
    
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "General Queries"
    
        Write-Verbose "Writing data to worksheet"

        #Headers
        $worksheet.Range("A1") = "Query"
        $worksheet.Range("B1") = "Status"
        $worksheet.Range("C1") = "Results"
        $worksheet.Range("D1") = "Duration"
        $worksheet.Range("A1:D1").Font.Bold=$True

        $i = 2

        #Writing summary List to worksheet
        foreach ($gen in $general) {
            $worksheet.Range("A$i") = $gen.Query
            $worksheet.Range("B$i") = $gen.Status
            $worksheet.Range("C$i") = $gen.Results.ToString("N0")
            $worksheet.Range("D$i") = $gen.Duration.ToString("hh\:mm\:ss\:fff")
            $i++
        }

        #Formatting
        $i--
        $worksheet.Range("A1:D$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("C1:C$i").HorizontalAlignment = -4108 #Align Center
        $worksheet.Range("A1:D$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "General Queries"'
    }

    Function WritePFExclusionsToExcel {
    
        Write-Verbose "Creating Predefined File Exclusions worksheet" 
    
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "Predefined File Exclusions"

        #Headers
        $worksheet.Range("A1") = "Reason"
        $worksheet.Range("B1") = "File"

        $i = 2

        #Writing Exlusions List to worksheet
        foreach ($Pfe in $predefinedExclusions) {
            $worksheet.Range("A$i") = $Pfe.Reason
            $worksheet.Range("B$i") = $Pfe.File
            $i++
        }

        #Formatting
        $i--
        $worksheet.Range("A1:B$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A1:B$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
                                   $worksheet.Range("A1:B$i"), $null, 
                                   [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "Predefined File Exclusions"'
    }

    Function WriteEngineConfigToExcel {
    
        Write-Verbose "Creating Engine Configuration worksheet" 
    
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Move([System.Type]::Missing, $workbook.Sheets.Item($workbook.Sheets.Count))
        $worksheet.Name = "Engine Configuration"

        #Headers
        $worksheet.Range("A1") = "Name"
        $worksheet.Range("B1") = "Value"

        $i = 2

        #Writing Exlusions List to worksheet
        foreach ($conf in $config.GetEnumerator()) {
            $worksheet.Range("A$i") = $conf.Name
            $worksheet.Range("B$i") = $conf.Value
            $i++
        }

        #Formatting
        $i--
        $worksheet.Range("A1:B$i").HorizontalAlignment = -4131 #Align Left
        $worksheet.Range("A1:B$i").Borders.LineStyle = 1
        [void]$worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
                                   $worksheet.Range("A1:B$i"), $null, 
                                   [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
        [void]$worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
        Write-Verbose 'Completed writing data to worksheet "Predefined File Exclusions"'
    }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Process

Process {

    #Display help if called
    if ($help -OR -NOT($logPath -XOR $scanId)) {
        Get-Help $MyInvocation.InvocationName -Full | Out-String
        exit
    }

    Write-Host "=========="
    $start = Get-Date
    Write-Host "Processing Started at $(Get-Date -Format "HH:mm:ss")"

    Write-Host "Loading log file"   
    if ($logPath) { 
        $lines = Get-Content -path $logPath
            Write-Host "Log file $logPath loaded"
    }
    else { 
        if($silentLogin) { $conn = New-SilentConnection $apiKey }
        else { $conn = New-Connection }     
        $lines = GetLogFile 
        Write-Host "Log file for Scan ID: $scanId loaded"
    }

    Write-Host "Parsing log file"
    ParseLogFile $lines
    Write-Host "Completed parsing"

    Write-Host "Creating Excel"
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Add()
    
    WriteDetailsToExcel
    WriteEngineConfigToExcel
    if ($details.PredefinedExclusions -gt 0) { WritePFExclusionsToExcel }
    WriteFilesToExcel
    WritePhasesToExcel
    WriteSummaryToExcel
    WriteGeneralToExcel

    $workbook.Worksheets.Item(1).Activate()

    Write-Host "Excel created"

    # Bring the Excel application to the front
    $excel.Visible = $true
    $excel.WindowState = [Microsoft.Office.Interop.Excel.XlWindowState]::xlMaximized

    # Release the COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    $end = Get-Date
    $runtime = (NEW-TIMESPAN –Start $start –End $end).ToString("hh\:mm\:ss")
    Write-Host "Processing Completed at $(Get-Date -Format "HH:mm:ss") with a runtime of $runtime"
    Write-Host "=========="
}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------