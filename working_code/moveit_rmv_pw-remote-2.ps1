###############################################################################################################
# Script to look for "MoveIT" service account ID's password in clear text inside files and delete those files #
# Author: Pablo Daniel García                                                                                 #
# Version: 2.3                                                                                                #
# Date: 26/Oct/2018                                                                                           #
###############################################################################################################

$t1 = Get-Date

##################################Functions###################################

Function now()
{
    return (Get-Date -Format "yyyy-MM-dd_HH_mm_ss")
}

Function today()
{
    return (Get-Date -Format "yyyy-MM-dd")
}

Function end($return_code)
{
    $t2 = Get-Date
    $delta_t = $t2 - $t1
    "$(now) - $($ThisServer) - INFO : Elapsed time $($delta_t.Hours) hours $($delta_t.Minutes) minutes $($delta_t.Seconds).$($delta_t.Milliseconds) seconds"|out-file -append $LogFile -encoding ASCII
    exit $return_code
}


###############################End of functions###############################

cls

$script_dir        = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$ThisServer        = $([System.Net.Dns]::GetHostByName((hostname)).HostName)
$maindir           = $script_dir
$LogFileDir        = "$($maindir)\Logs"
$LogFile           = "$($LogFileDir)\$(now)-$($ThisServer)-local.log"
$Logfileslist      = @()
$Outputfilesdir    = "$($maindir)\Outputfiles"
$np_servers_list   = "$($OutputFilesDir)\GRID-servers_not_processed_list.txt"
$RemoteLogFile     = $null
$inputdir          = "c:\temp"
$Search_pattern    = "xomdmz\xsmoveit"
$servers_list_file = "$($script_dir)\GRID-servers_list-test-2.txt"  # input file used for testing purposes
#$servers_list_file = "$($script_dir)\GRID-servers_list.txt"
$domain            = "xomdmz.net"
$psv               = $PSVersionTable.PSVersion.ToString()

# Create LogFileDir and LogFile

if(! (test-path $LogFileDir))
{
    new-item -path $LogFileDir -itemtype Directory -force -erroraction silentlycontinue > $null
}
	

if (! (test-path $LogFile))
{
    new-item -path $LogFile -itemtype File -force -erroraction silentlycontinue > $null
}

if(! (test-path $Outputfilesdir))
{
    new-item -path $Outputfilesdir -itemtype Directory -force -erroraction silentlycontinue > $null
}

if(Test-Path $np_servers_list)  # if the file with the list of servers not processed exists, delete it
{
    Remove-Item $np_servers_list -Force -erroraction silentlycontinue > $null
}
New-Item -Path $np_servers_list -ItemType File -force -erroraction silentlycontinue > $null  # create new file to hold list of servers not processed

"$(now) - $($ThisServer) - INFO : MoveIT clear text password checker powershell main script started"|out-file -append $LogFile -encoding ASCII
"$(now) - $($ThisServer) - INFO : Log file directory '$($LogFileDir)' exists"|out-file -append $LogFile -encoding ASCII
"$(now) - $($ThisServer) - INFO : Log file '$($LogFile)' exists"|out-file -append $LogFile -encoding ASCII
"$(now) - $($ThisServer) - INFO : Main script is running on server '$($ThisServer)'"|out-file -append $LogFile -encoding ASCII
"$(now) - $($ThisServer) - INFO : The installed Powershell version is ($($psv))"|out-file -append $LogFile -encoding ASCII


# Read list of servers from text file
if(Test-Path $servers_list_file)
{
    $servers_list = Get-Content $servers_list_file -ErrorAction SilentlyContinue
    "$(now) - $($ThisServer) - INFO : Servers list file '$($servers_list_file)' read"|out-file -append $LogFile -encoding ASCII
}
else  # if file not found, end the script
{
    "$(now) - $($ThisServer) - ERRO : Servers list file '$($servers_list_file)' does not exist. Ending script"|out-file -append $LogFile -encoding ASCII
    end(1)
}


##############################################################################
#          Script command block that is executed on remote servers           #
##############################################################################

$command =
{
$t1 = Get-Date

##################################Functions###################################

Function now()
{
    return (Get-Date -Format "yyyy-MM-dd_HH_mm_ss")
}

Function today()
{
    return (Get-Date -Format "yyyy-MM-dd")
}

Function end($return_code)
{
    $t2 = Get-Date
    $delta_t = $t2 - $t1
    "$(now) - $($ThisServer) - INFO : Elapsed time $($delta_t.Hours) hours $($delta_t.Minutes) minutes $($delta_t.Seconds).$($delta_t.Milliseconds) seconds"
    "$($return_code)"
    exit $return_code
}

Function check_mod($mod)
{
    if (Get-Module | Where-Object { $_.name -eq $mod })
    {
        return ($true)
    }
    else
    {
        return ($false)
    }
}

Function mod_avail($mod)
{
    if (Get-Module -ListAvailable | Where-Object { $_.name -eq $mod })
    {
        return ($true)
    }
    else
    {
        return ($false)
    }
}

###############################End of functions###############################

$ThisServer        = $([System.Net.Dns]::GetHostByName((hostname)).HostName)
$inputdir          = "c:\temp"
$Search_pattern    = "xomdmz\xsmoveit"
$psv               = $PSVersionTable.PSVersion.ToString()


"$(now) - $($ThisServer) - INFO : MoveIT clear text password checker powershell remote script started"
"$(now) - $($ThisServer) - INFO : Script is running on server '$($ThisServer)'"
"$(now) - $($ThisServer) - INFO : The installed Powershell version is ($($psv))"

# check if needed module to handle zip files is available in the installed version of Powershell
# If it is, check if is loaded. If it is not loaded, load it
$module  = "Microsoft.PowerShell.Archive"
$mod_loaded = $false
if (mod_avail($module))
{
    "$(now) - $($ThisServer) - INFO : Module '$($module)' is available in the installed version ($($psv)) of PowerShell. Attempting to load it"
    Import-Module $module
    if(check_mod($module))
    {
        "$(now) - $($ThisServer) - INFO : Module '$($module)' loaded"
        $mod_loaded = $true
        $zip_uncompression_mode = "1"
    }
    else
    {
        "$(now) - $($ThisServer) - ERRO : Module '$($module)' not loaded"
        "$(now) - $($ThisServer) - ERRO : $($Error[0].Exception.GetType().FullName)"
        "$(now) - $($ThisServer) - ERRO : $($Error[0] | Format-List * -Force)"
    }
}
else
{
    "$(now) - $($ThisServer) - ERRO : Module '$($module)' is not available in the installed version ($($psv)) of PowerShell"
}

if(!($mod_loaded))  # if required module is not loaded
{
    $zip_uncompression_mode = "2"
    "$(now) - $($ThisServer) - INFO : Checking if .Net assembly to handle zip files is loaded"
    if ([appdomain]::CurrentDomain.GetAssemblies() | ?{$_ -ilike "System.IO.Compression.FileSystem"})
    {
        "$(now) - $($ThisServer) - INFO : .Net assembly to handle zip files was already loaded"
    }
    else
    {
        "$(now) - $($ThisServer) - INFO : .Net assembly to handle zip files is not loaded. Attempting to load it"
        Add-Type -assembly "System.IO.Compression.FileSystem"
        if ([appdomain]::CurrentDomain.GetAssemblies() | ?{$_ -ilike "System.IO.Compression.FileSystem"})
        {
            "$(now) - $($ThisServer) - INFO : .Net assembly to handle zip files is now loaded"
        }
        else
        {
            "$(now) - $($ThisServer) - INFO : .Net assembly to handle zip files could not be loaded. Aborting script execution"
            end(1)
        }
    }
}


"$(now) - $($ThisServer) - INFO : Directory where search will be performed is '$($inputdir)'"

# Find installation source directory and delete it
"$(now) - $($ThisServer) - INFO : Checking if 'MoveIT' is installed on this server ($($ThisServer))"
$inst_prod_classes = Get-ChildItem -Path Registry::HKLM\SOFTWARE\Classes\Installer\Products -Force -Recurse
$source_dir = $null
$source_file = $null
foreach($r in $inst_prod_classes)
{
    if($r.Property -eq "ProductName")
    {
        $prod_name = (Get-ItemProperty -Path Registry::$r -Name ProductName).ProductName
        if($prod_name -like "*moveit*")
        {
            $class = $r.PSChildName
            $prod_ver = (Get-ItemProperty -Path Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$class\InstallProperties -Name DisplayVersion).DisplayVersion
            $source_dir = ((Get-ItemProperty -Path Registry::$r\SourceList).LastUsedSource).Split(";",3)[2]
            $source_file = (Get-ItemProperty -Path Registry::$r\SourceList).PackageName
            $source = "$($source_dir)$($source_file)"
            "$(now) - $($ThisServer) - INFO : MoveIT is published as '$($prod_name) $($prod_ver)'"
            "$(now) - $($ThisServer) - INFO : MoveIT source file is '$($source)'"
        }
    }
}
if($source_dir)
{
    "$(now) - $($ThisServer) - INFO : Attempting to delete directory '$($source_dir)'"
    if(Test-Path $source_dir)
    {
        Remove-Item -Recurse -Force $source_dir -erroraction silentlycontinue > $null  # delete directory and all its contents - comment for testing purposes
        if(Test-Path $source_dir)
        {
            "$(now) - $($ThisServer) - ERRO : Directory '$($source_dir)' could not be deleted"
        }
        else
        {
            "$(now) - $($ThisServer) - INFO : Deleted directory '$($source_dir)'"
        }
    }
    else
    {
        "$(now) - $($ThisServer) - INFO : Directory $($source_dir) does not exist"
    }
}
else
{
    "$(now) - $($ThisServer) - INFO : MoveIT is not installed on this server"
}


# Build list of zip files in all directory structure inside $inputdir
"$(now) - $($ThisServer) - INFO : Looking for zip files in directory '$($inputdir)'"
$zip_files_list = @()
$zip_files_list = Get-ChildItem -Path $inputdir -Recurse -File -Force -Include *.zip
$zip_files_to_delete = @()
if($zip_files_list.Count -gt 0)
{
    "$(now) - $($ThisServer) - INFO : Beginning search for '$($Search_pattern)' in $($zip_files_list.Count) zip file(s) found in directory '$($inputdir)'"
    foreach($zf in $zip_files_list)
    {
        "$(now) - $($ThisServer) - INFO : Searching for '$($Search_pattern)' in '$($zf)' zip file"
        $temp_dir = "$($inputdir)\$(now) - $($ThisServer)"
        new-item -path $temp_dir -itemtype Directory -force -erroraction silentlycontinue > $null
        Switch -Exact ($zip_uncompression_mode)
        {
            "1"
            {
                Expand-Archive -Path $zf.FullName -DestinationPath $temp_dir -Force  # extract files from current zip file to temporary directory
                ;Break
            }

            "2"
            {
                [System.IO.Compression.ZipFile]::ExtractToDirectory($zf.FullName, $temp_dir)
                ;Break
            }
        }
        $temp_files_list = Get-ChildItem -Path $temp_dir -Recurse -File -Force  # build list of temporary files for current zip file

        # Process zip files list
        $found = $false
        foreach($file in $temp_files_list)
        {
            $search = Select-String -Pattern $Search_pattern -Path $file -SimpleMatch
            if($search -ne $null)  # if there is a match break the loop and add the zip file to the list of zip files to be deleted
            {
                $f = $file.FullName
                $f = $f -ireplace [regex]::Escape($temp_dir), $zf
                "$(now) - $($ThisServer) - WARN : Found one occurrence of '$($Search_pattern)' in line $($search.LineNumber[0]) of file '$($f)'"
                "$(now) - $($ThisServer) - WARN : Flagging '$($zf)' for deletion and stopping the search in this file"
                $found = $true
                $zip_files_to_delete += $zf
                break
            }
        }
        if(! $found)
        {
            "$(now) - $($ThisServer) - INFO : '$($Search_pattern)' not found on any files inside '$($zf.FullName)'"
        }
        "$(now) - $($ThisServer) - INFO : attempting to delete temporary directory '$($temp_dir)' used for analysis of contents of '$($zf.FullName)'"
        Remove-Item -Recurse -Force $temp_dir -erroraction silentlycontinue > $null  # delete temporary directory where zip file's contents have been extracted for analysis
        if(Test-Path $temp_dir)
        {
            "$(now) - $($ThisServer) - ERRO : temporary directory '$($temp_dir)' could not be deleted. Delete it manually"
        }
        else
        {
            "$(now) - $($ThisServer) - INFO : temporary directory '$($temp_dir)' deleted"
        }
    }
    "$(now) - $($ThisServer) - INFO : Ended search for '$($Search_pattern)' in $($zip_files_list.Count) zip file(s) found in directory '$($inputdir)'"
    if($zip_files_to_delete.Count -gt '0')  # if there were any matches on any files inside any of the zip files
    {
        [array]::Sort($zip_files_to_delete) > $null
        "$(now) - $($ThisServer) - INFO : Attempting to delete $($zip_files_to_delete.Count) zip file(s) flagged to be erased"
        foreach($rf in $zip_files_to_delete)
        {
            Remove-Item -Force $rf -erroraction silentlycontinue > $null  # delete each zip file flagged for erasure - comment for testing purposes
            if(Test-Path $rf)
            {
                "$(now) - $($ThisServer) - ERRO : File '$($rf)' could not be deleted"
            }
            else
            {
                "$(now) - $($ThisServer) - INFO : Deleted file '$($rf)'"
            }
        }
    }
    else
    {
        "$(now) - $($ThisServer) - INFO : '$($Search_pattern)' not found on any of the $($zip_files_list.Count) zip file(s) inside '$($inputdir)'"
    }
}
else
{
    "$(now) - $($ThisServer) - INFO : No zip files found in directory '$($inputdir)'"
}

# Build list of files in all directory structure inside $inputdir excluding zip files (already processed)
$files_list = Get-ChildItem -Path $inputdir -Recurse -File -Force -Exclude *.zip

"$(now) - $($ThisServer) - INFO : Searching '$($Search_pattern)' in $($files_list.Count) files excluding zip files (already processed)"

# Process files list
$found = @()
$found_dir = @()
$i = 0
foreach($file in $files_list)
{
    $search = Select-String -Pattern $Search_pattern -Path $file -SimpleMatch
    if($search -ne $null)  # if there is a match add it to the found list
    {
        $found += $search
        $found_dir += $file.DirectoryName  # add the corresponding directory to the list of directories where file with match was found
        for($j=0; $j -lt $search.Count; $j++)  # process multiple matches in the same file
        {
            "$(now) - $($ThisServer) - WARN : $($found[$i].Path);$($found[$i].LineNumber)"  # log only file path and line number of each match excluding line itself to avoid logging the sensitive data connected to the search
            $i ++
        }
    }
}

if($found.Count -gt '0')  # if there were matches
{
    $found_file=@()
    foreach($f in $found)
    {
        $found_file += $f.Path  # get the file name from the search string
    }
    [array]::Sort($found_file) > $null  # sort list of found files using .Net function -faster and more efficient than powershell sort-
    $found_file = $found_file | Get-Unique -AsString  # remove duplicates
    [array]::Sort($found_dir) > $null  # sort list of directory where files were found using .Net function -faster and more efficient than powershell sort-
    $found_dir = $found_dir | Get-Unique -AsString  # remove duplicates
    "$(now) - $($ThisServer) - INFO : $($found.Count) occurrence(s) of '$($Search_pattern)' found in $($found_file.Count) file(s) in $($found_dir.Count) directory(ies) excluding zip files (processed separatedly)"
    
    # delete files where string pattern was found and log that
    foreach($rfile in $found_file)
    {
        "$(now) - $($ThisServer) - INFO : Attempting to delete file '$($rfile)'"
        Remove-Item -Force $rfile -erroraction silentlycontinue > $null  # delete file where string pattern has been found - comment for testing purposes
        if(Test-Path $rfile)
        {
            "$(now) - $($ThisServer) - ERRO : File '$($rfile)' could not be deleted"
        }
        else
        {
            "$(now) - $($ThisServer) - INFO : Deleted file '$($rfile)'"
        }
    }
}
else  # pattern not found in any files
{
    "$(now) - $($ThisServer) - INFO : '$($Search_pattern)' was not found in any uncompressed file in '$($inputdir)' (zip files processed separatedly)"
}
end(0)
}
##############################################################################
#                         End of script command block                        #
##############################################################################


# Start jobs
$job_id = @()
Foreach($server in $servers_list)
{
    $fqdn = "$($server).$($domain)"
    $jid = (Invoke-Command -ComputerName $fqdn -JobName $server -AsJob -ScriptBlock $command).Id
    "$(now) - $($ThisServer) - INFO : Sent job for remote server '$($fqdn)' to execute with ID $($jid)" |out-file -append $LogFile -encoding ASCII
    $job_id += $jid
    $RemoteLogFile = "$($LogFileDir)\$(now)-$($ThisServer)-remote($($server)).log"
    new-item -path $RemoteLogFile -itemtype File -force -erroraction silentlycontinue > $null
    if(Test-Path($RemoteLogFile))
    {
        "$(now) - $($ThisServer) - INFO : Log file '$($RemoteLogFile)' for remote system '$($fqdn)' created"|out-file -append $LogFile -encoding ASCII
        $Logfileslist += [PSCustomObject] @{
            server = $server;
            file = $RemoteLogFile;
        }

    }
    else
    {
        "$(now) - $($ThisServer) - ERRO : Log file '$($RemoteLogFile)' for remote system '$($fqdn)' could not be created. Output from the corresponding will not be locally logged"|out-file -append $LogFile -encoding ASCII
        $Logfileslist += [PSCustomObject] @{
            server = $server;
            file = $null;  # if log file could not be created, assign null value to file field. Pending the addition of logic to handle this circumstance when retrieving jobs' output
        }
    }

}

# Scan jobs until they:
#1) finish and retrieve their data and remove them from memory, or
#2) fail and remove them from memory, delete their corresponding log file and add the server on which it failied to execute to the list of non processed servers

$jobs_running = $true
do
{
    $job = Get-Job
    $jobs_running = $false
    foreach($j in $job)
        {
            if($j.Id -in $job_id)
            {
                switch -Wildcard ($j.State)
                {
                    "Completed"  # job completed
                    {
                        "$(now) - $($ThisServer) - INFO : Job on remote server '$($j.Name).$($domain)' with ID $($j.Id) ended"|out-file -append $LogFile -encoding ASCII
                        if($j.HasMoreData)
                        {
                            $job_output = Receive-Job $j
                            switch -Exact ($job_output[($job_output.Count)-1])
                            {
                            "0"  #job was completely executed
                                {
                                    "$(now) - $($ThisServer) - INFO : Job on remote server '$($j.Name).$($domain)' with ID $($j.Id) was completely executed"|out-file -append $LogFile -encoding ASCII
                                    ; Break
                                }
                            "1"  #job was not completely executed - no search performed because there was no way to process zip files available in the installed version of Powershell
                                {
                                    "$(now) - $($ThisServer) - CRIT : Job on remote server '$($j.Name).$($domain)' with ID $($j.Id) was not completely executed - no search was performed because there was no way to process zip files available in the installed version of Powershell"|out-file -append $LogFile -encoding ASCII
                                    ; Break
                                }
                            }
                            $job_output[0..($job_output.Count-2)] |out-file -append $Logfileslist[$Logfileslist.server.IndexOf($j.Name)].file -encoding ASCII
                        }
                        else
                        {
                            "$(now) - $($ThisServer) - ERRO : Data from job on remote server '$($j.Name).$($domain)' with ID $($j.Id) could not be retrieved"|out-file -append $LogFile -encoding ASCII
                        }
                        Remove-Job $j  # clear job from memory
                        ; Break
                    }

                    "Failed"  # job failed
                    {
                        "$(now) - $($ThisServer) - ERRO : Job on remote server '$($j.Name).$($domain)' with ID $($j.Id) failed"|out-file -append $LogFile -encoding ASCII
                        $j.Name|out-file -append $np_servers_list -encoding ASCII
                        $lf_to_del = $Logfileslist[$Logfileslist.server.IndexOf($j.Name)].file
                        "$(now) - $($ThisServer) - INFO : Attempting to delete log file '$($lf_to_del)'"|out-file -append $LogFile -encoding ASCII
                        Remove-Item $lf_to_del -Force -ErrorAction silentlycontinue > $null
                        if(Test-Path $lf_to_del)
                        {
                            "$(now) - $($ThisServer) - WARN : Log file '$($lf_to_del)' could not be deleted"|out-file -append $LogFile -encoding ASCII
                        }
                        else
                        {
                            "$(now) - $($ThisServer) - INFO : Log file '$($lf_to_del)' deleted"|out-file -append $LogFile -encoding ASCII
                        }
                        Remove-Job $j  # clear job from memory
                        ; Break
                    }

                    default
                    {
                        $jobs_running = $true
                    }
                }
            }
        }
}
while($jobs_running)

end(0)