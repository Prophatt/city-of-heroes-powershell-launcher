#support for secure web connections
Add-Type -AssemblyName System.Web
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#local folder of script
$wd = get-item -path ".\" -verbose

$webclient = New-Object System.Net.WebClient 
#array for manifests that use multiple URL elements to perform load balancing
$site_collection = New-Object System.collections.arraylist($null)

function start-pause{
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}
#remove files that are 0 in size, some manifests use those to make updates etc
function remove-0files{
    $0_files = Get-ChildItem -Recurse -Path $wd
    foreach ($1_file in $0_Files){
    if ($1_file.length -eq 0){
        Remove-Item $1_file.fullname -force -Exclude "errorlog*"
        }
}
}
#errorlog function
function start-logging{
    if ( -Not (Test-Path "$wd\errorlog.dat" ))
    {
        $error_log = new-item -Path "$wd\errorlog.dat" -ItemType File
    }
    else {
        Remove-Item -Path "$wd\errorlog.dat" -force 
        $error_log = new-item -Path "$wd\errorlog.dat" -ItemType File
    }
    Start-Transcript -Path $error_log -Force -NoClobber -Append
}

function exit-launcher{
    Write-Host ""
    Write-Host "Quitting"
    Write-Host ""
    Stop-Transcript
    Exit
}

function write-splash {
Write-Host ""
Write-Host "Welcome to the Powershell console launcher for City of Heroes"
Write-Host ""
}
#retrieve initial manifest, should be consolidated into manifest compare
function get-manifestupdate {
    parameter([string]$single_manifest)    
    Clear-Host
    $check_manifest = Import-Csv -Path "$wd\manifests\manifest_address.csv"
    foreach ($line in $check_manifest){
        if ($single_manifest -eq $line.name){
            $manifest_name = $line.name
            $xml_path = "$wd\manifests\$manifest_name"
            $webclient.DownloadFile($line.address,$xml_path)
        Write-Host ""
        Write-Host "The Manifest has been downloaded, files are now being downloaded."
        write-host $xml_path
        Write-Host "Please press any Key to continue"
        start-pause
        break
        }
    }
}
#download files from manifest
function get-files {
    #timer to compare download total time with other launcher download times
    $sw = [Diagnostics.Stopwatch]::StartNew()
    parameter([string]$xml_path)
    Clear-Host
    get-manifestupdate($single_manifest)
    $xml = [xml](Get-Content $xml_path)
    $xml_info = $xml.manifest.filelist.file
    #read files and download if either non-existent or md5 is different
    foreach($item in $xml_info){
        Write-Host ""
        $name = $item.name
        Write-Host "File is located in <city of heroes base directory>/$name"
        [int]$url_count = $item.url.count
        $file_hash = $item.md5
        $file_url = $item.url
        Write-Host "The file hash is $file_hash"
        #for manifests that use multiple url elements for load balancing
        if ($url_count -gt 1){
            foreach ($file in $item.url){
                $site_collection.add($file) | Out-Null
            }
            $random_sites = Get-Random -inputobject $site_collection
            $file_url = $random_sites
            $error.clear()
            $webclient.openread($file_url) | Out-Null
           
            if (!$error){
                write-host "THE WEBSITE IS OK"
            }else{
                do{
                Write-Host "THERE IS AN ERROR WITH THE WEBSITE $file_url, PICKING NEW WEBSITE"
                $error.clear()
                $site_collection.remove($file_url) | Out-Null
                $random_sites = Get-Random -inputobject $site_collection
                $file_url = $random_sites
                $webclient.openread($file_url) | Out-Null
                
                }until (!$error)
                write-host "THE WEBSITE IS OK"
            }       
            Write-Host "The download link is $file_url"
        }else{Write-Host "The download link is $file_url"}
        #clear site array for next element
        $site_collection.clear()

        $fileName = $file_url.Substring($file_url.LastIndexOf('/')+1)
        #split manifest file name in case it includes a folder, create folder if necessary
        if ($name -match "/")
            {
        $split_folder = $name -split "/"
        $internal_folder = $split_folder[0]
            if ( -Not (Test-Path "$wd\$internal_folder" ))
                {
            New-Item -Path "$wd\$internal_folder" -ItemType Directory | Out-Null
                }
           #download file/create folder if non existent
            If ( -Not (Test-Path "$wd\$internal_folder\$fileName" )){
                $error.Clear()
                $webclient.DownloadFile($file_url,"$wd\$internal_folder\$fileName")
                if ($error){
                    $error.clear()
                    write-host "Error on download. Attempting old download method."
                    write-host "The url is $file_url"
                    $file_link = [System.Web.HttpUtility]::UrlDecode($file_url)
                    $fileName = $file_link.Substring($file_link.LastIndexOf('/')+1)
                    $downloadfile = invoke-webrequest -uri $file_link
                    $file_contents = $downloadfile.Content
                    [io.file]::WriteAllBytes("$wd\$internal_folder\$filename",$file_contents)
                }
                #check md5 and download new file if md5 is different
            }Else {
                $local_hash = Get-ChildItem -path "$wd\$internal_folder\$filename" | Get-FileHash -Algorithm MD5
                if ($file_hash -eq $local_hash.hash){
                    write-host "Hash of local file and server file are the same, no additional download needed."
                }Else{
                    write-host "Hash of the local file and server file are different. Redownloading."
                    Remove-Item -Path "$wd\$internal_folder\$filename" -Force
                    $error.Clear()
                    $webclient.DownloadFile($file_url,"$wd\$internal_folder\$fileName")
                    if ($error){
                        $error.clear()
                        write-host "Error on download. Attempting old download method."
                        write-host "The url is $file_url"
                        $file_link = [System.Web.HttpUtility]::UrlDecode($file_url)
                        $fileName = $file_link.Substring($file_link.LastIndexOf('/')+1)
                        $downloadfile = invoke-webrequest -uri $file_link
                        $file_contents = $downloadfile.Content
                        [io.file]::WriteAllBytes("$wd\$internal_folder\$filename",$file_contents)
                    }
                }
            }
        }
            Else{
                #download file/create folder if non existent
            If ( -Not (Test-Path "$wd\$fileName" )){
                $error.Clear()
                $webclient.DownloadFile($file_url,"$wd\$fileName")
                if ($error){
                    $error.clear()
                    write-host "Error on download. Attempting old download method."
                    write-host "The url is $file_url"
                    $file_link = [System.Web.HttpUtility]::UrlDecode($file_url)
                    $fileName = $file_link.Substring($file_link.LastIndexOf('/')+1)
                    $downloadfile = invoke-webrequest -uri $file_link
                    $file_contents = $downloadfile.Content
                    [io.file]::WriteAllBytes("$wd\$filename",$file_contents)
                }
            }Else {      
                #check md5 and download new file if md5 is different note this should all be consolidated into above logic 
                $local_hash = Get-ChildItem -path "$wd\$filename" | Get-FileHash -Algorithm MD5
            if ($file_hash -eq $local_hash.hash){
                write-host "Hash of local file and server file are the same, no additional download needed."
            }Else{
                write-host "Hash of the local file and server file are different. Redownloading."
                Remove-Item -Path "$wd\$filename" -Force
                $error.Clear()
                $webclient.DownloadFile($file_url,"$wd\$fileName")
                if ($error){
                    $error.clear()
                    write-host "Error on download. Attempting old download method."
                    write-host "The url is $file_url"
                    $file_link = [System.Web.HttpUtility]::UrlDecode($file_url)
                    $fileName = $file_link.Substring($file_link.LastIndexOf('/')+1)
                    $downloadfile = invoke-webrequest -uri $file_link
                    $file_contents = $downloadfile.Content
                    [io.file]::WriteallBytes("$wd\$filename",$file_contents)
                }
                }
                }
                }
        }
        
        remove-0files     
        
        write-host ""
        write-host ""
        $sw.Stop()
        write-host "Time taken to download - "$sw.Elapsed
        Write-Host ""
        write-host "Files have been downloaded, press a key to continue"
        start-pause
}
#validate files
function get-validation {
    parameter([string]$single_manifest)
    Clear-Host
    $xml_path = "$manifest_directory\$single_manifest"
    get-files([string]$xml_path)
    #write dynamic launcher menu to screen after validation
    get-launcher("$single_manifest")
}
#add manifest to launcher
function get-newmanifest{
    Clear-Host
    write-splash
    if (-not (Test-Path "$wd\manifests\manifest_address.csv")){
        New-Item -Path "$wd\manifests\manifest_address.csv" -ItemType file | Out-Null
        $manifest_paths = Get-ChildItem -Path "$wd\manifests\manifest_address.csv"
        Add-Content $manifest_paths "name,address"
        }else {
            $manifest_paths = Get-ChildItem -Path "$wd\manifests\manifest_address.csv"
        }
    $manifest_name = Read-Host "Please give the manifest a name, ie Rebirth, COXG, Victory, Homecoming etc"
    $manifest_name = $manifest_name + ".xml"
    $new_manifest = Read-Host "Please place the manifest url here. Tip - you can right click to paste"
#check for duplicate manifest name
$manifest_csv = Get-Content -Path $manifest_paths | Select-String -Pattern $manifest_name
    if ($manifest_csv -match $manifest_name){
        Write-Host ""
        Write-Host "There is already a manifest with this name."
        Write-Host "Please press any key to continue"
        start-pause
        get-manifests
    }else{Add-Content $manifest_paths "$manifest_name,$new_manifest"}

    $xml_path = "$manifest_directory\$manifest_name"
    $webclient.DownloadFile($new_manifest,$xml_path)
Write-Host ""
Write-Host "The Manifest has been downloaded, files are now being downloaded."
write-host $xml_path
#begin game download
get-files($xml_path)
get-manifests
}
#check for manifest update, download new manifest to temp file, if md5 of manifests do not match delete original and rename temp to original name
function compare-manifest {
    Parameter([string]$single_manifest)  
    Clear-Host
    write-splash
    $check_manifest = Import-Csv -Path "$wd\manifests\manifest_address.csv"
        foreach ($line in $check_manifest){
          
            if ($single_manifest -match $line.name){
                $manifest_name = $line.name
                $xml_path = "$wd\manifests\$manifest_name"
                $temp_manifest_name = "temp_$manifest_name"
                $temp_xml_path = "$wd\manifests\$temp_manifest_name"
                write-host "Checking for manifest update."
                $webclient.DownloadFile($line.address,$temp_xml_path)
                $local_hash = Get-ChildItem -path $xml_path | Get-FileHash -Algorithm MD5
                $temp_hash = Get-ChildItem -path $temp_xml_path | Get-FileHash -Algorithm MD5
                                if ($temp_hash.hash -eq $local_hash.hash){
                                  remove-item -force $temp_xml_path
                        get-launcher ($single_manifest)
                                }Else{
                                    write-host "Hash of local manifest and server manifest are different."
                remove-item -force $xml_path
                rename-item -path $temp_xml_path -newname $manifest_name
                        write-host "New Manifest has been installed."
                        write-host "Please press any key to check for updates."
                start-pause
                get-files($xml_path)
  start-pause
    }
}
} 
}



#manifest list for initial screen load so user can choose between different servers
function get-manifests{
Clear-Host
$manifest_directory = "$wd\manifests"
if ( -Not (Test-Path $manifest_directory ))
{
New-Item -Path $manifest_directory -ItemType Directory | Out-Null
}
$found_manifests = Get-ChildItem -filter "*.xml" -path $wd"\manifests"

write-splash
Write-Host "The following manifests have been found."
Write-Host ""
#write manifest to screen
$manifest_array = @{}
#set menu number so that users do not have zero as a menu choice
$manifest_count = 1
foreach ($single_manifest in $found_manifests){
[string]$manifest_num = $manifest_count
[string]$manifest_menu = $manifest_num + '   ' + $single_manifest
$manifest_array.Add($manifest_count, $single_manifest)
Write-Host $manifest_menu
$manifest_count = $manifest_count + 1
}
write-host ""
write-host "A   Add New Manifest"
#write-host "V   Revalidate Manifest Files"
write-host "Q   Quit"
Write-Host ""
$main_choice = Read-Host "Please make a selection."
switch($main_choice){
    "Q" {
        exit-launcher
    }
    "A" {
        get-newmanifest
    }
    default{
        #adjust array numbering so that it matches on screen numbering
        [int]$corrected_choice = $main_choice
       $single_manifest = $manifest_array[$corrected_choice]
       compare-manifest($single_manifest)
        get-launcher($single_manifest)
    }

}
}


#launcher menu for chosen manifest
function get-launcher{

    Parameter([string]$single_manifest)
Clear-Host
write-splash

$launcher_array = @{}
$launch_manifest = [xml] (Get-Content -Path "$wd\manifests\$single_manifest")
$launcher_menu = $launch_manifest.manifest.profiles.launch
$count = 1
#get launcher options from chosen manifest and dynamically write to screen
foreach($item in $launcher_menu){
    [string]$menu_num = $count
    [string]$menu_string = $menu_num + '   ' + $item.innertext
    Write-Host $menu_string
    
    [string]$launch_command = $item.exec + ' ' + $item.params
    $launcher_array.Add($count, $launch_command)
    [string]$farm_params = $item.params
    $count = $count + 1
}
#extra static options
write-host ""
write-host "F   Start with Farming Options"
write-host "D   Start with Console for Troubleshooting Purposes"
write-host "C   Check for Manifest Update"
write-host "V   Revalidate Only Files"
write-host "R   Return to Previous Screen"
write-host "Q   Quit"
Write-Host ""
$main_choice = Read-Host "Please make a selection."
switch($main_choice){
    "F" {
        write-host ""
        [int]$server_choice = Read-Host "Please choose a server from the above list to start with client farming options."
        #[int]$corrected_choice = $main_choice
        [string]$command_string = $launcher_array[$server_choice]
        $launch_farmer = $command_string + " -stopinactivedisplay 1"
        #write-host $launch_farmer
        #start-pause
        cmd /c $launch_farmer
        get-launcher($single_manifest)
        }
    "D" {
        write-host ""
        [int]$server_choice = Read-Host "Please choose a server from the above list to start with console."
        #[int]$corrected_choice = $main_choice
        [string]$command_string = $launcher_array[$server_choice]
        $launch_farmer = $command_string + " -console"
        #write-host $launch_farmer
        #start-pause
        cmd /c $launch_farmer
        get-launcher($single_manifest)
        }
    "C" {
        compare-manifest($single_manifest)
        }
    "Q" {
        exit-launcher
        }
    "R" {
        #return to server manifest selection screen
        get-manifests
        }
    "V" {
        get-validation($single_manifest)
        }   
    default{
        [int]$corrected_choice = $main_choice
        $command_string = $launcher_array[$corrected_choice]
        cmd /c $command_string
        get-launcher($single_manifest)
        }
    }
}

start-logging
get-manifests