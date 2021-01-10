#######################################################################################
# Import Module
Import-Module $PSScriptRoot\utils\Write-Menu.psm1
Import-Module $PSScriptRoot\utils\Show-Pause.psm1

#######################################################################################
# Select Font
# ref:https://stackoverflow.com/questions/58855377/add-menu-options-in-a-running-powershell-script
Write-Host "Show the list of Nerd-Font latest version"
$ReleasePage = "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
$Json = Invoke-WebRequest $ReleasePage | ConvertFrom-Json
$LatestVersion = $Json[0].tag_name

$Font_Name_Extend = Write-Menu -Title 'Select one Nerd-Font' -Entries @($Json[0].assets.name)
$Font_Name = $Font_Name_Extend.Split(".")[0]
Write-Host $Font_Name_Extend

#######################################################################################
# Download Sources from github
$DownloadUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/$LatestVersion/$Font_Name_Extend"

#check whether *.zip file exist
if (-not(Test-Path $PSScriptRoot\$Font_Name_Extend)) {
    Write-Host Downloading latest release
    Invoke-WebRequest $DownloadUrl -Out $PSScriptRoot\$Font_Name_Extend
}
else {
    Write-Host Already Downloads the Font
}
#check whether font.zip already expan or not
if (-not(Test-Path $PSScriptRoot\$Font_Name\)) {
    Write-Host Extracting release files
    Expand-Archive -LiteralPath $PSScriptRoot\$Font_Name_Extend -DestinationPath $PSScriptRoot\$Font_Name
}
else {
    Write-Host Already expand archive
}

########################################################################################
# Insatll Font 
#Reference:
#https://github.com/mikeTWC1984/pwshise/blob/ee3209eac079e4b0adc40c569f25ec77a75a6ad3/fonts/FontInstaller.ps1
#https://gist.github.com/anthonyeden/0088b07de8951403a643a8485af2709b
#https://richardspowershellblog.wordpress.com/2008/03/20/special-folders/
#https://jordanmalcolm.com/deploying-windows-10-fonts-at-scale/
if ($IsWindows) {
    $Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
    # Check if already installed 
    # '*.ttf', '*.ttc', '*.otf'
    Get-ChildItem -Path $PSScriptRoot\$Font_Name -Include '*.otf' -Recurse | ForEach-Object {
        if (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
        #if (-not(Test-Path "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\$($_.Name)")) {   
            Write-Host Installing font  $($_.BaseName)
            
            # Install font for current user
            #$Destination.CopyHere($_.FullName, 0x14)

            # Install for all user
            Copy-Item $_.FullName "C:\Windows\Fonts"
            New-ItemProperty -Name $_.BaseName -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" -PropertyType string -Value $_.name         
        }
        else {
            Write-Host $($_.Name) already installed
        }
    }
}
#if($IsLinux){

#}
#if($IsMacOS){}

#
# Clean
Write-Host Cleanup $Font_Name folder and zip
Remove-Item $PSScriptRoot\$Font_Name -Recurse -ErrorAction SilentlyContinue 
Remove-Item $PSScriptRoot\$Font_Name_Extend