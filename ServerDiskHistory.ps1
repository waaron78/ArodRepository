#get disk usage information and export it to a CSV file for trend reporting


Param(
    [string[]]$ServerName = $env:COMPUTERNAME
)

#path to CSV file is hard coded because I always want to use this file
$CSV = "c:\work\serverdiskhistory.csv"

#initialize an empty array
$data = @()

#define a hashtable of parameters to splat to Get-CimInstance
$cimParams = @{
    Classname   = "Win32_LogicalDisk"
    Filter      = "drivetype = 3"
    ErrorAction = "Stop"
}


foreach ($i in $ServerName) {

Write-Host "Getting disk information from $i." -ForegroundColor Cyan
    #update the hashtable on the fly
    $cimParams.Computername = $i
    Try {
        $disks = Get-CimInstance @cimparams

        $data += $disks |
            Select-Object @{Name = "Computername"; Expression = {$_.SystemName}},
        DeviceID, Size, FreeSpace,
        @{Name = "PctFree"; Expression = { ($_.FreeSpace / $_.size) * 100}},
        @{Name = "Date"; Expression = {Get-Date}}
    } #try
    Catch {
        Write-Warning "Failed to get disk data from $($ServerName.toUpper()). $($_.Exception.message)"
    } #catch
} #foreach

#only export if there is something in $data
if ($data) {
    $data | Export-Csv -Path $csv -Append -NoTypeInformation
    Write-Host "Disk report complete. See $CSV." -ForegroundColor Green
}
else {
    Write-Host "No disk data found." -ForegroundColor Yellow
}

#sample usage
# .\serverdiskhistory..ps1 -ServerName Srv1,Srv2,Srv3,Srv4
