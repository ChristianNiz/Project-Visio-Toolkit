# Microsoft Project-Visio Toolkit v2.1
# Function: To uninstall MS Office 2016 "Click-to-Run" Version and Install Compatiable Office/Project/Visio


Function uninstall_microsoft_office
{
Robocopy "\\PATH\toolkit\Uninstall tool" C:\temp\Office_uninstall /MIR /R:0 /w:0
Start-Process -FilePath C:\Temp\Office_uninstall\O15CTRRemove.diagcab 
main_menu
}

Function install_Microsoft_Office_2016_x64
{
Robocopy "\\PATH\toolkit\Microsoft_Office_2016_Pro_x64" C:\temp\Office_Installation /MIR /R:0 /w:0 /ETA
Start-Process C:\Temp\Office_Installation\setup.exe
main_menu
}

Function install_Microsoft_Office_2016_x32 
{
Robocopy "\\PATH\toolkit\Microsoft_Office_2016_Pro_x32" C:\temp\Office_Installation /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation\setup.exe 
main_menu
}

Function install_Microsoft_Visio_PRO_2016_x64 
{
Robocopy "\\PATH\toolkit\VisioPro2016x64" C:\temp\Office_Installation_Visio /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Visio\setup.exe 
main_menu
}

Function install_Microsoft_Visio_PRO_2016_x32 
{
Robocopy "\\PATH\toolkit\VisioPro2016x32" C:\temp\Office_Installation_Visio /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Visio\setup.exe 
main_menu
}

Function install_Microsoft_Visio_STD_2016_x64 
{
Robocopy "\\PATH\toolkit\VisioStd2016x64" C:\temp\Office_Installation_Visio /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Visio\setup.exe 
main_menu
}

Function install_Microsoft_Visio_STD_2016_x32 
{
Robocopy "PATH\toolkit\VisioStd2016x32" C:\temp\Office_Installation_Visio /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Visio\setup.exe 
main_menu
}

Function install_Microsoft_Project_PRO_2016_x64 
{
Robocopy "PATH\toolkit\Project_Pro_2016_x64" C:\temp\Office_Installation_Project /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Project\setup.exe 
main_menu
}

Function install_Microsoft_Project_PRO_2016_x32 
{
Robocopy "\\PATH\toolkit\Project_Pro_2016_x32" C:\temp\Office_Installation_Project /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Project\setup.exe 
main_menu
}

Function install_Microsoft_Project_STD_2016_x64 
{
Robocopy "\\PATH\toolkit\Project_std_2016_x64" C:\temp\Office_Installation_Project /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Project\setup.exe 
main_menu
}
Function install_Microsoft_Project_STD_2016_x32 
{
Robocopy "\\PATH\toolkit\Project_std_2016_x32" C:\temp\Office_Installation_Project /MIR /R:0 /w:0
Start-Process C:\Temp\Office_Installation_Project\setup.exe 
main_menu
}

Function KMS_Activation

{

    $url = "PATH/scripts/Microsoft_KMS-MAK_activation.bat" 
    $path = "C:\temp\Microsoft_KMS-MAK_activation.bat" 
   # param([string]$url, [string]$path) 
      
    if(!(Split-Path -parent $path) -or !(Test-Path -pathType Container (Split-Path -parent $path))) { 
      $path = Join-Path $pwd (Split-Path -leaf $path) 
    } 
      
    "Downloading [$url]`nSaving at [$path]" 
    $client = new-object System.Net.WebClient 
    $client.DownloadFile($url, $path) 
    #$client.DownloadData($url, $path) 
      
    $path
	
Start-Process cmd.exe -Verb RunAs "/c C:\temp\Microsoft_KMS-MAK_activation.bat"
}

function main_menu{
write-host
write-host
write-host		  "############################################################" 
write-host		  "# Please make sure you're connected to the Internal Network #"
write-host		  "#                                                          #"
write-host		  "#                 Created By CNIZ                          #"
write-host		  "#                                                          #"
write-host		  "#    Microsoft Office-Project-Visio-Toolkit v2.1           #"
write-host		  "#                                                          #"
write-host		  "#   Note: This is both 64bit and 32bit compatible          #"
write-host		  "#                                                          #"
write-host		  "#        *Uninstall Microsoft Office                       #"
write-host		  "#        *Install Office 2016 (64 or 32bit)                #"
write-host		  "#        *Install Visio and Project (64 or 32bit)          #"
write-host		  "############################################################"
write-host
write-host "1: Uninstall Microsoft Office"
write-host
write-host "2: Install Microsoft Office Professional 2016 x64"
write-host "3: Install Microsoft Office Professional 2016 x32"
write-host
write-host "4: Install Microsoft Visio 2016 Professional x64"
write-host "5: Install Microsoft Visio 2016 Professional x32"
write-host
write-host "6: Install Microsoft Visio 2016 Standard x64"
write-host "7: Install Microsoft Visio 2016 Standard x32"
write-host
write-host "8: Install Microsoft Project 2016 Professional x64"
write-host "9: Install Microsoft Project 2016 Professional x32"
write-host
write-host "10: Install Microsoft Project 2016 Standard x64"
write-host "11: Install Microsoft Project 2016 Standard x32"
Write-host
Write-host "12: KMS Activate Microsoft Office"
Write-host 
write-host "13: Exit"`n

	$input = Read-Host "Select Option(s) From Menu"
	switch ($input)
	{
		'1' { uninstall_microsoft_office } #uninstall current microsoft office
		'2' { install_Microsoft_Office_2016_x64 } #install current microsoft office x64
		'3' { install_microsoft_office_2016_x32 } #install current microsoft office x32
		'4' { install_Microsoft_Visio_PRO_2016_x64  } #Install MS Visio Professional 2016 32bit
		'5' { install_Microsoft_Visio_PRO_2016_x32  } #Install MS Visio Professional 2016 64bit
		'6' { Install_microsoft_Visio_STD_2016_x64  } #Install MS Visio Standard 2016 32bit
		'7' { Install_microsoft_Visio_STD_2016_x32  } #Install MS Visio Standard 2016 64bit
		'8' { install_Microsoft_Project_PRO_2016_x64  } #Install MS Project Professional 2016 32bit
		'9' { Install_microsoft_Project_PRO_2016_x32  } #Install MS Project Professional 2016 64bit
		'10' { Install_microsoft_Project_STD_2016_x64  } #Install MS Project Standard 2016 32bit
		'11' { Install_microsoft_Project_STD_2016_x32  } #Install MS Project Standard 2016 64bit
		'12' {KMS_Activation} #KMS_Activation of Microsoft Office 2016
		'13' {stop-transcript Exit} #Exit and Clean up
	}
}

main_menu
