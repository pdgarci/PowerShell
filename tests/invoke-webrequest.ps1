clear
$name="directory_name"    #subdirctory within the web domain from where files will be got
$url_dommain = "https://web_domain"    #web domain to be targetted by the GET requests
$url_prefix = "$($url_domain)/$($name)"
$output_file_path = "$($env:HOMEDRIVE)$($env:HOMEPATH)\Downloads\$name"
$digits = ("000","00","0","")    #assuming there are at most 10000 files with sequence numbers [0..9999]
if (-not (Test-Path $output_file_path))
{
    mkdir $output_file_path
}
$suffixes = ".jpg",".mp4"  #files suffixes (types) to fetch 
$files_count = 100       #number of files sequentially named & numbered to fetch
foreach ($suffix in $suffixes)
{
    for ($i=1; $i -le $files_count; $i++)
    {
        $zeros = $digits[[Math]::Truncate([Math]::Log10($i))]
        try {Invoke-WebRequest -Uri $($url_prefix+$zeros+$i+$suff) -OutFile $($output_file_path+'\'+$name+'_'+$zeros+$i+$suffix) -Headers @{"Cache-Control"="no-cache"}}
        catch [System.Net.WebException],[System.IO.IOException] {
            "Unable to download $($name)_$($zeros)$($i)$($suff) from '$($url_domain)'"
        }
    }
}