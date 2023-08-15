clear
$script_dir        = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$script_name       = $MyInvocation.MyCommand.Name.Split(".",2)[0]   # it is assumed that script file name is "script_name.*"
$ThisComputer      = $([System.Net.Dns]::GetHostByName((hostname)).HostName)
$maindir           = $script_dir
$LogFileDir        = "$($maindir)\$($script_name)\Logs"
#$LogFile           = "$($LogFileDir)\$(now)-$($ThisComputer)-local.log"
$psv               = $PSVersionTable.PSVersion.ToString()

$encoded_file = [System.IO.File]::ReadAllBytes("$($script_dir)\1.txt")
$decoded_file = [System.Convert]::FromBase64CharArray($encoded_file, 0, $encoded_file.Length)
[System.IO.File]::WriteAllBytes("$($script_dir)\2.jpg", $decoded_file)
