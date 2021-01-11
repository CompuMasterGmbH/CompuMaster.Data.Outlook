::REM -UpdateNuGetExecutable not required since it's updated by VS.NET mechanisms
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& 'CompuMaster.Data.Outlook\_CreateNewNuGetPackage\DoNotModify\New-NuGetPackage.ps1' -ProjectFilePath '.\CompuMaster.Data.Outlook\CompuMaster.Data.Outlook.vbproj' -verbose -NoPrompt -DoNotUpdateNuSpecFile -PushPackageToNuGetGallery"
pause