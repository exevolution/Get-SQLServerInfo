Clear-Host
# Remove previous run
If ($FailTable) {Remove-Variable FailTable -Force | Out-Null}
If ($SuccessTable) {Remove-Variable SuccessTable -Force | Out-Null}

# Define Important Stuff
$ScriptRoot = $PSScriptRoot

# Get all server names from the saved txt file
$Servers = @(Import-Csv -LiteralPath "$ScriptRoot\SQLserverlist.csv")

# Import product definitions
$ProductArray = @(Import-Csv -LiteralPath "$ScriptRoot\defProducts.csv")

# Create data tables
$SuccessTable = New-Object -TypeName System.Data.DataTable "SQL Server Info"
$FailTable = New-Object -TypeName System.Data.DataTable "Failed Servers"

# Define Columns
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn Server,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn OperatingSystemSKU,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn ServicePack,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn OperatingSystem,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn SQLVersion,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn SQLEdition,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn SQLInstallPath,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn RAM,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn CPU,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn AllDrives,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD1,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD2,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD3,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD4,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD5,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD6,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD7,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD8,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD9,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD10,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD11,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD12,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD13,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD14,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD15,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD16,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD17,([string])))
$SuccessTable.Columns.Add((New-Object System.Data.DataColumn HDD18,([string])))


$FailTable.Columns.Add((New-Object System.Data.DataColumn Server,([string])))
$FailTable.Columns.Add((New-Object System.Data.DataColumn ThrownException,([string])))

# Loop through each server 
ForEach ($Server in $Servers)
{
    Try
    {
        Test-NetConnection $Server.HostName -InformationLevel Quiet -ErrorAction Stop | Out-Null
        $OS = Get-WmiObject -ClassName Win32_OperatingSystem -ComputerName $Server.HostName -Property * -ErrorAction Stop | Select-Object *
        #$OS = Get-CimInstance -ClassName CIM_OperatingSystem -ComputerName $Server.HostName -Property *
        $CPU = Get-WmiObject -ClassName Win32_Processor -ComputerName $Server.HostName -Property * -ErrorAction Stop | Select-Object *
        #$CPU = Get-CimInstance -ClassName CIM_Processor -ComputerName $Server.HostName -Property *
        $HDD = Get-WmiObject -ClassName Win32_LogicalDisk -ComputerName $Server.HostName -Property * -ErrorAction Stop | Select-Object * | Where-Object {$_.DriveType -eq 3} | Sort-Object -Property DeviceID
        #$HDD = Get-CimInstance -ClassName CIM_LogicalDisk -ComputerName $Server.HostName -Property * | Where-Object {$_.DriveType -eq 3} | Sort-Object -Property DeviceID
        $RAM = [Convert]::ToInt64(($OS | Select-Object -ExpandProperty TotalVisibleMemorySize), 10)
        $RAM = [Math]::Round($RAM / 1MB, 0)
        $ComputerManagementNamespace = (Get-WmiObject -ComputerName $Server.HostName -Namespace "root\microsoft\sqlserver" -Class "__NAMESPACE" | Where-Object {$_.Name -like "ComputerManagement*"} | Select-Object Name | Sort-Object Name -Descending | Select-Object -First 1).Name
        If ($ComputerManagementNamespace -eq $Null)
        {
            Write-Error "ComputerManagement namespace not found"
        }
        Else
        {
            $ComputerManagementNamespace = "root\microsoft\sqlserver\" + $ComputerManagementNamespace
        } 
        $SQLEdition = Get-WmiObject -ComputerName $Server.HostName -Namespace $ComputerManagementNamespace -Class "SqlServiceAdvancedProperty" | Where-Object {$_.ServiceName -eq "MSSQLSERVER" -and $_.PropertyName -eq "SKUNAME"} |
        Select-Object @{Name = "PropertyValue"; Expression = {If ($_.PropertyValueType -eq 0) {$_.PropertyStrValue} Else {$_.PropertyNumValue}}}
        $SQLPaths = (Get-WmiObject -ComputerName $Server.HostName -Class Win32_Service | Where-Object {$_.PathName -like '*sqlservr.exe*'} | Select-Object -ExpandProperty PathName).TrimEnd(" ") | Out-String

    }
    Catch [System.Management.Automation.ActionPreferenceStopException]
    {
        Write-Error "$($Server.HostName) encountered an error"
        $FailRow = $FailTable.NewRow()
        $FailRow.Server = $Server.HostName
        $FailRow.ThrownException = "[$($Error[0].Exception.GetType().FullName)] $($Error[0].Exception.Message)"
        $FailTable.Rows.Add($FailRow)
        Continue
    }
    Catch [System.Runtime.InteropServices.COMException]
    {
        Write-Error "$($Server.HostName): $($Error[0].Exception.Message)"
        $FailRow = $FailTable.NewRow()
        $FailRow.Server = $Server.HostName
        $FailRow.ThrownException = "[$($Error[0].Exception.GetType().FullName)] $($Error[0].Exception.Message)"
        $FailTable.Rows.Add($FailRow)
        Continue
    }
    Catch [System.Management.Automation.ItemNotFoundException]
    {
        Write-Warning "$($Error[0].Exception.Message)"
        $FailRow = $FailTable.NewRow()
        $FailRow.Server = $Server.HostName
        $FailRow.ThrownException = "[$($Error[0].Exception.GetType().FullName)] $($Error[0].Exception.Message)"
        $FailTable.Rows.Add($FailRow)
    }
    Catch
    {
        Write-Warning "$($Error[0].Exception.Message)"
        $FailRow = $FailTable.NewRow()
        $FailRow.Server = $Server.HostName
        $FailRow.ThrownException = "[$($Error[0].Exception.GetType().FullName)] $($Error[0].Exception.Message)"
        $FailTable.Rows.Add($FailRow)
    }

    If ($SQLVersion)
    {
        Remove-Variable SQLVersion -Force
    }
    $SQLInstall = @()
    ForEach ($SQLPath in $SQLPaths)
    {
        $SQLPath = $SQLPath.Substring(1,($SQLPath.IndexOf('"',2))-1)
        $SQLPath = $SQLPath -replace ':','$'
        $PathCheck = "\\$($Server.HostName)\$SQLPath"
        If (Test-Path $PathCheck)
        {
            $SQLVersion = @(Get-Item -LiteralPath "$PathCheck" | Select-Object -ExpandProperty VersionInfo).ProductVersion
        }
        If ($SQLVersion)
        {
            $SQLInstall += $PathCheck
            "SQL Instance located at $SQLInstall"
        }
        Else
        {
            "SQL not detected on $($Server.HostName)"
            $SQLInstall = "Not detected as an installed service"
        }
    }
    Switch -Wildcard ($SQLVersion)
    {
        # SQL Server 2016
        "13.1.4*" {$VersionName = "SQL Server 2016 SP1 $SQLVersion"}
        "13.0.4*" {$VersionName = "SQL Server 2016 SP1 $SQLVersion"}
        "13.0.2*" {$VersionName = "SQL Server 2016 $SQLVersion"}
        "13.0.1*" {$VersionName = "SQL Server 2016 $SQLVersion"}
        # SQL Server 2014
        "12.2.5*" {$VersionName = "SQL Server 2014 SP2 $SQLVersion"}
        "12.0.5*" {$VersionName = "SQL Server 2014 SP2 $SQLVersion"}
        "12.1.4*" {$VersionName = "SQL Server 2014 SP1 $SQLVersion"}
        "12.0.4*" {$VersionName = "SQL Server 2014 SP1 $SQLVersion"}
        "12.0.2*" {$VersionName = "SQL Server 2014 $SQLVersion"}
        # SQL Server 2012
        "11.3.6*" {$VersionName = "SQL Server 2012 SP3 $SQLVersion"}
        "11.0.6*" {$VersionName = "SQL Server 2012 SP3 $SQLVersion"}
        "11.2.5*" {$VersionName = "SQL Server 2012 SP2 $SQLVersion"}
        "11.0.5*" {$VersionName = "SQL Server 2012 SP2 $SQLVersion"}
        "11.1.3*" {$VersionName = "SQL Server 2012 SP1 $SQLVersion"}
        "11.0.3*" {$VersionName = "SQL Server 2012 SP1 $SQLVersion"}
        "11.0.2*" {$VersionName = "SQL Server 2012 $SQLVersion"}
        # SQL Server 2008 R2
        "10.53.6*" {$VersionName = "SQL Server 2008 R2 SP3 $SQLVersion"}
        "10.50.6*" {$VersionName = "SQL Server 2008 R2 SP3 $SQLVersion"}
        "10.52.4*" {$VersionName = "SQL Server 2008 R2 SP2 $SQLVersion"}
        "10.50.4*" {$VersionName = "SQL Server 2008 R2 SP2 $SQLVersion"}
        "10.51.2*" {$VersionName = "SQL Server 2008 R2 SP1 $SQLVersion"}
        "10.50.2*" {$VersionName = "SQL Server 2008 R2 SP1 $SQLVersion"}
        "10.50.1*" {$VersionName = "SQL Server 2008 R2 $SQLVersion"}
        # SQL Server 2008
        "10.4.6*" {$VersionName = "SQL Server 2008 SP4 $SQLVersion"}
        "10.0.6*" {$VersionName = "SQL Server 2008 SP4 $SQLVersion"}
        "10.3.5*" {$VersionName = "SQL Server 2008 SP3 $SQLVersion"}
        "10.0.5*" {$VersionName = "SQL Server 2008 SP3 $SQLVersion"}
        "10.2.4*" {$VersionName = "SQL Server 2008 SP2 $SQLVersion"}
        "10.0.4*" {$VersionName = "SQL Server 2008 SP2 $SQLVersion"}
        "10.1.2*" {$VersionName = "SQL Server 2008 SP1 $SQLVersion"}
        "10.0.2*" {$VersionName = "SQL Server 2008 SP1 $SQLVersion"}
        "10.0.1*" {$VersionName = "SQL Server 2008 $SQLVersion"}
        # SQL Server 2005
        "9*" {$VersionName = "SQL Server 2005 $SQLVersion"}
        Default {$VersionName = "Unknown Version $SQLVersion"}
    }

    $SuccessRow = $SuccessTable.NewRow()
    $SuccessRow.Server = $Server.HostName
    $SuccessRow.OperatingSystem = $OS.Caption
    $SuccessRow.OperatingSystemSKU = ($ProductArray | Where-Object {$_.Value -eq $OS.OperatingSystemSKU})."SKU Name"
    $SuccessRow.ServicePack = $OS.CSDVersion
    $SuccessRow.SQLVersion = $VersionName
    $SuccessRow.SQLEdition = $SQLEdition.PropertyValue
    $SuccessRow.SQLInstallPath = ($SQLInstall | Out-String)
    $SuccessRow.RAM = "$([Convert]::ToString($RAM))"
    $SuccessRow.CPU = ($CPU | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
    $i = 0
    $AllDrives = @()
    ForEach ($Drive in $HDD)
    {
        $i++
        $DriveOutput = "{0} {1}GB" -F $Drive.DeviceID,[Math]::Round($Drive.Size/1GB, 0)
        $CurrentDrive = "HDD" + $i
        $SuccessRow.$CurrentDrive = $DriveOutput
        $AllDrives += $DriveOutput
    }
    $SuccessRow.AllDrives = $AllDrives | Out-String
    $SuccessTable.Rows.Add($SuccessRow)
}

Try
{
    $SuccessTable | Export-Csv -LiteralPath "$ScriptRoot\Get-SQLServerInfo.csv" -NoTypeInformation
    $FailTable | Export-Csv -LiteralPath "$ScriptRoot\Get-SQLServerInfo-failures.csv" -NoTypeInformation
    Invoke-Item "$ScriptRoot\Get-SQLServerInfo.csv"
    Invoke-Item "$ScriptRoot\Get-SQLServerInfo-failures.csv"
}
Catch
{
    $FailRow = $FailTable.NewRow()
    $FailRow.Server = ""
    $FailRow.ThrownException = "[$($Error[0].Exception.GetType().FullName)] $($Error[0].Exception.Message)"
    $FailTable.Rows.Add($FailRow)
}
