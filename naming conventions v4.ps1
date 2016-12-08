################################
########## PARAMETERS ##########
################################  
$ScanFolder = "C:\MyPackages"
$NamingConventions = "C:\MyPackages\JamieThomson.csv"
$ReportFolder = "C:\MyPackages"


#################################################
########## DO NOT EDIT BELOW THIS LINE ##########
#################################################
clear
Write-Host "=========================================================================================="
Write-Host "==                                  Used parameters                                     =="
Write-Host "=========================================================================================="
Write-Host "Folder to scan         :" $ScanFolder
Write-Host "Naming Conventions CSV :" $NamingConventions
Write-Host "Report folder          :" $ReportFolder
Write-Host "=========================================================================================="

$startTime = [System.DateTime]::Now

# Include functions from secondairy file
. "$PSScriptRoot\functions.ps1"



#############################################
########## LOAD NAMING CONVENTIONS ##########
#############################################
# Load CSV with naming conventions into hash table
$hashtablePrefix = loadNamingConventions($NamingConventions)
 
# Show hash table ordered for debug purposes
# $hashtablePrefix.GetEnumerator() | sort -Property value | Format-Table -Auto



############################################
########## CREATE ERROR LOG TABLE ##########
############################################
#Create Table object for errors
$errorTable = New-Object system.Data.DataTable “ErrorList”

#Define Columns
$col1 = New-Object system.Data.DataColumn Solution,([string])
$col2 = New-Object system.Data.DataColumn Project,([string])
$col3 = New-Object system.Data.DataColumn Package,([string])
$col4 = New-Object system.Data.DataColumn Path,([string])
$col5 = New-Object system.Data.DataColumn Name,([string])
$col6 = New-Object system.Data.DataColumn Prefix,([string])
$col7 = New-Object system.Data.DataColumn Error,([string])

#Add the Columns
$errorTable.columns.add($col1)
$errorTable.columns.add($col2)
$errorTable.columns.add($col3)
$errorTable.columns.add($col4)
$errorTable.columns.add($col5)
$errorTable.columns.add($col6)
$errorTable.columns.add($col7)



##############################################
########## CHECK NAMING CONVENTIONS ##########
##############################################
# Counter for number of package
$PackageCount = 0

# Loop through all packages in all subfolders, but exclude the OBJ folder 
Get-ChildItem $ScanFolder -Filter *.dtsx -Recurse | ? {$_.FullName -notmatch "\\obj\\"} | Foreach-Object {
    [xml]$pkg = Get-Content $_.FullName
    $TasksContainers = $pkg.Executable.Executables.Executable
    loopTasksAndComponents $TasksContainers $_.FullName

    # Logging to screen to keep you entertained and busy
    # But it will cost a couple of extra seconds
    if ($PackageCount -ne 0 -and $PackageCount % 900 -eq 0)
    {
        # new line after 90 dots
        Write-Host ""
    }
    if ($PackageCount % 10 -eq 0)
    {
        # write dot for each tenth package
        Write-Host -NoNewline "."
    }
    $PackageCount = $PackageCount + 1
}
Write-Host ""
Write-Host "Packages analyzed   :" $PackageCount 

#Display the table
#$errorTable | format-table -AutoSize 



#NOTE: Now you can also export this table to a CSV file as shown below.
$errorTable | Where-Object {$_.Prefix -ne "?"} | export-csv (Join-Path -Path $ReportFolder -ChildPath "Errors.csv") -noType

#NOTE: Now you can also export this table to a CSV file as shown below.
$errorTable | Where-Object {$_.Prefix -eq "?"} | export-csv (Join-Path -Path $ReportFolder -ChildPath "Unknown.csv") -noType

# Log errorcount to screen
Write-Host "Wrong prefixes found:" ($errorTable | Where-Object {$_.Prefix -ne "?"}).Count
Write-Host "Unknown types found :" ($errorTable | Where-Object {$_.Prefix -eq "?"}).Count



$TimeSpan = NEW-TIMESPAN –Start $StartTime –End ([System.DateTime]::Now)
Write-Host "Duration:" $TimeSpan.Minutes "minutes," $TimeSpan.Seconds "seconds"
Write-Host "Reports saved in" $ReportFolder