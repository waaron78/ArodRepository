#  Script Name : Automation_Monitor_Disk_Space.ps1
#
#
# Purpose : To notify Engineers on disk space status of different disk drives by html report based email
#
# 
#
#             

#Setting up variables
$BASE_DIR=(Resolve-Path .\).Path

$LOG_FILE=$BASE_DIR + "\daily_disk_space_monitor.log"


$head = @"
<style>
h1, h5, th { text-align: center; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #dae5f4; }
tr:nth-child(odd) { background: #b8d1f3; }
</style>
"@
$report=$BASE_DIR + "\daily_diskSpace_status.html"

$xml_config=$BASE_DIR + "\configuration.xml"
[xml]$xml_content=Get-Content $xml_config


write-output "$(get-date) :INFO Staring the script Execution " | out-file $LOG_FILE -Append -Force;  

$diskspace_error_threshold = [int32] $xml_content.DAILY_REPORT.MONITOR_THRESHOLD.DISKSPACE_ERROR

$diskspace_warning_threshold = [int32] $xml_content.DAILY_REPORT.MONITOR_THRESHOLD.DISKSPACE_WARNING



#==================================Preparing  Report=======================================================

	Try{
	
		foreach ($entity in $xml_content.DAILY_REPORT.MONITOR_SERVERS ){ 
				 $server = $entity.SERVER

                $body="<h3>Daily Summary</h3>`n<h3>Updated: on $(Get-Date)</h3>"
               
			   Get-WmiObject  -Query "select * from Win32_LogicalDisk where drivetype=3" -ComputerName $server  |
						Select-Object SystemName,DeviceId, @{Name="Size";Expression={[math]::Round($_.Size/1GB,2)}}`
                                                    ,@{Name="FreeSpace";Expression={[math]::Round($_.FreeSpace/1GB,2)}} `
                                                    ,@{Name="Occupied"; Expression= {[math]::Round(100 - ( [double]$_.FreeSpace / [double]$_.Size ) * 100)}} |
								Export-Csv disk_space.csv -NoTypeInformation
				
				}
				
				$csv_content = Import-CSV 'disk_space.csv'
				
				




		$body="<h3>Servers in Error Status</h3>`n<h3>Updated: on $(Get-Date)</h3>"
		$csv_content | Where-Object {$_.Occupied -ge $diskspace_error_threshold}  | ConvertTo-Html -property SystemName,DeviceId,Size,FreeSpace,Occupied `
		-Head $head -Body $body | Out-File $report 
					
		 $body="<h3>Servers in Warning Status</h3>`n<h3>Updated: on $(Get-Date)</h3>"
		 $csv_content | Where-Object {$_.Occupied -lt $diskspace_error_threshold -and $_.Occupied -gt $diskspace_warning_threshold }  |
 ConvertTo-Html -property SystemName,DeviceId,Size,FreeSpace,Occupied `
        -Body $body | Out-File $report -Append


		 $body="<h3>Servers in Good Status</h3>`n<h3>Updated: on $(Get-Date)</h3>"
		 $csv_content | Where-Object {$_.Occupied -lt $diskspace_warning_threshold}  | ConvertTo-Html -property SystemName,DeviceId,Size,FreeSpace,Occupied `
        -Body $body | Out-File $report -Append
				
	



}Catch{

		$ErrorMessage = $_.Exception.Message
		write-output "$(get-date) :ERROR Something went Wrong ErrorMessage :  $ErrorMessage " | out-file $LOG_FILE -Append -Force; 
	} 

write-output "$(get-date) :INFO Report Preparation Execution Over" | out-file $LOG_FILE -Append -Force; 

#==================================Report Prepared=======================================================






#==================================SENDING EMAIL=======================================================
Try{
	$SMTPServer = $xml_content.DAILY_REPORT.EMAIL.SMTP
	$Username = $xml_content.DAILY_REPORT.EMAIL.SMTP_USERNAME
	$Password = $xml_content.DAILY_REPORT.EMAIL.SMTP_PASSWORD

	$message = New-Object System.Net.Mail.MailMessage
	$message.subject = $xml_content.DAILY_REPORT.EMAIL.SUBJECT
	$message.body = Get-Content $report
	$message.to.add( $xml_content.DAILY_REPORT.EMAIL.TO )
	$message.cc.add($xml_content.DAILY_REPORT.EMAIL.CC)
	$message.IsBodyHtml = $True
	$message.from = $xml_content.DAILY_REPORT.EMAIL.FROM


	$smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
	$smtp.EnableSSL = $true
	$smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
	$smtp.send($message)
	write-output "$(get-date) :INFO Email Sent" | out-file $LOG_FILE -Append -Force; 
	
} catch {

	$ErrorMessage = $_.Exception.Message
	write-output "$(get-date) :ERROR Something went Wrong.  ErrorMessage :  $ErrorMessage " | out-file $LOG_FILE -Append -Force; 
		
} finally{
		write-output "$(get-date) : Script Execution Completed " | out-file $LOG_FILE -Append -Force; 
	
}
#==================================SENDING EMAIL COMPLETED=======================================================



