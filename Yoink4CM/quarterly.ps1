## If using Notepad, save file to desktop and copy / paste back tp C:\Program Files\Yoink Software\Yoink4CM\quarterly.ps1
## If using Notepad++, it will prompt to run as administrator to enable saving in the above location.

& winget download VideoLAN.VLC -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Notepad++.Notepad++ -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Dell.CommandUpdate -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Python.Python.3.13 -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Corel.WinZip -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download 7zip.7zip -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download RARLab.WinRAR -d $Tempfolder --accept-package-agreements --accept-source-agreements