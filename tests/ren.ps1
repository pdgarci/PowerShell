clear
cd "$($env:USERPROFILE)\Downloads"
dir | rename-item -NewName {$_.name -replace 'string_to_delete',””}