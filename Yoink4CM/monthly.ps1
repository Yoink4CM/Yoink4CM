## If using Notepad, save file to desktop and copy / paste back tp C:\Program Files\Yoink Software\Yoink4CM\monthly.ps1
## If using Notepad++, it will prompt to run as administrator to enable saving in the above location.

& winget download Google.Chrome -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Mozilla.Firefox -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Zoom.Zoom -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Adobe.Acrobat.Reader.64-bit -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Microsoft.Teams -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Microsoft.VCRedist.2015+.x86 -d $Tempfolder --accept-package-agreements --accept-source-agreements
& winget download Microsoft.VCRedist.2015+.x64 -d $Tempfolder --accept-package-agreements --accept-source-agreements