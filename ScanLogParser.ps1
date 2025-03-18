#------------------------------------------------------------------------------------------------------------------------------------------
#region Help

<#
.Synopsis
Opens an Excel file with the Parsed results of a Checkmarx SAST Scan Log 

.Description
Takes a scan log file as an input and opens an excel with parsed details from the log file
Has tabs for General Details, Engine Configuration, Predefined File Exclusions, Phases, Files, Results Summary and General Queries 
Excel created is not saved and must be manually saved if required

Usage
Help
    .\ScanLogParser.ps1 -help [<CommonParameters>]
    
Parse Log File
    .\ScanLogParser.ps1 -logPath <string> [<CommonParameters>] 

.Notes
Version:     1.0
Date:        18/03/2025
Written by:  Michael Fowler
Contact:     michael.fowler@checkmarx.com

Change Log
Version    Detail
-----------------
1.0        Original version
  
.PARAMETER help
Display help

.PARAMETER logPath
The file path for the Scan Log to be processes

#>

#endregion
#------------------------------------------------------------------------------------------------------------------------------------------
#region Parameters

[CmdletBinding(DefaultParametersetName='Parse')] 
Param (

    [Parameter(ParameterSetName='Help',Mandatory=$false, HelpMessage="Display help")]
    [switch]$help,

    [Parameter(ParameterSetName='Parse',Mandatory=$true, HelpMessage="Enter Full path for scan log")]
    [string]$logPath
)

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Data Structures
    
    $summary = [System.Collections.Generic.List[ResultsSummary]]::New()
    $general = [System.Collections.Generic.List[GeneralQuery]]::New()
    $files = [System.Collections.Generic.List[File]]::New()
    $predefinedExclusions = [System.Collections.Generic.List[Exclusion]]::New()
    $excludeFiles = [System.Collections.Generic.List[String]]::New()
    $phases = [System.Collections.Generic.List[Phase]]::New()
    $details = [LogDetails]::new()
    $config = @{}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Results Summary Class

class ResultsSummary {
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

    [String]$Query
    [String]$Severity
    [String]$Status
    [Int]$Results
    [Nullable[TimeSpan]]$Duration
    [String]$Cwe

    #endregion    
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Constructors
        
        ResultsSummary ([String] $line) {
            $this.Query = ($line.Substring(0,88)).Trim()
            $this.Severity = ($line.Substring(98,13)).Trim()
            $this.Status = ($line.Substring(111,9)).Trim()
            $this.Results = ($line.Substring(129,7)).Trim()
            $this.Duration = [TimeSpan]::ParseExact($line.Substring(147,12), "hh\:mm\:ss\.fff", $null) 
            $this.Cwe = ($line.Substring(162,12)).Trim()
        }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region General Query Class

class GeneralQuery {

    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

    [String]$Query
    [String]$Status
    [Int]$Results
    [Nullable[TimeSpan]]$Duration

    #endregion    
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Constructors

        GeneralQuery ([String] $line) {
            $this.Query = ($line.Substring(0,80)).Trim()
            $this.Status = ($line.Substring(80,17)).Trim()
            $this.Results = ($line.Substring(100,14)).Trim()
            $this.Duration = [TimeSpan]::ParseExact($line.Substring(114,12), "hh\:mm\:ss\.fff", $null)
        }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------

}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Log Details Class

class LogDetails {
    #------------------------------------------------------------------------------------------------------------------------------------------------
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
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Constructors

        LogDetails() { }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------

}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region File Class

class File {
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

        [Nullable[datetime]]$Start
        [Nullable[datetime]]$End
        [Nullable[TimeSpan]]$Runtime
        [String]$FileName

    #endregion    
    #------------------------------------------------------------------------------------------------------------------------------------------------
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
    #------------------------------------------------------------------------------------------------------------------------------------------------

}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Phase Class

class Phase {
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

        [Nullable[datetime]]$Start
        [Nullable[datetime]]$End
        [Nullable[TimeSpan]]$Runtime
        [String]$PhaseName

    #endregion    
    #------------------------------------------------------------------------------------------------------------------------------------------------
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
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Exclusion Class

class Exclusion {
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Variables

        [String]$Reason
        [String]$File

    #endregion    
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Constructors

        Exclusion([String]$reason, [String]$file) {
            $this.Reason = $reason
            $this.File = $file
        }

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------

}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Functions

Function ParseLogFile {
   
    Write-Verbose "Loading log file"   

    $lines = Get-Content -path $logPath
    
    Write-Verbose "Log file $logPath loaded"  

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

    Write-Verbose "Writing data to worksheet"
    
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Details 
    
    #Header
    $worksheet.Range("A1:B1").Merge()
    $worksheet.Range("A1:B1").HorizontalAlignment = -4108
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
   
    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    #region Parsing Summary 
    
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

    #endregion
    #------------------------------------------------------------------------------------------------------------------------------------------------
    
    #Auto-fit columns
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
    Write-Verbose 'Completed writing data to worksheet "Details"'
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:F$i"))
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$worksheet.Range("A1:D$i"))
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
                               $worksheet.Range("A1:B$i"), $null, 
                               [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
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
    $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, 
                               $worksheet.Range("A1:B$i"), $null, 
                               [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $worksheet.UsedRange.Cells.EntireColumn.AutoFit()
    
    Write-Verbose 'Completed writing data to worksheet "Predefined File Exclusions"'
}

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------
#region Main

#Display help if called
if ($help) {
    Get-Help $MyInvocation.InvocationName -Full | Out-String
    exit
}

Write-Host "=========="
Write-Host "Processing Started at $(Get-Date -Format "HH:mm:ss dd/MM/yyyy")"

Write-Host "Parsing log file"
ParseLogFile
Write-Host "Completed parsing"

Write-Host "Creating Excel"
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()

WriteDetailsToExcel | Out-Null
WriteEngineConfigToExcel | Out-Null
if ($details.PredefinedExclusions -gt 0) { WritePFExclusionsToExcel | Out-Null }
WriteFilesToExcel | Out-Null
WritePhasesToExcel | Out-Null
WriteSummaryToExcel | Out-Null
WriteGeneralToExcel | Out-Null

$workbook.Worksheets.Item(1).Activate()

Write-Host "Excel created"

$excel.Visible = $true

Write-Host "Processing Completed at $(Get-Date -Format "HH:mm:ss dd/MM/yyyy")"
Write-Host "=========="

#endregion
#----------------------------------------------------------------------------------------------------------------------------------------------------