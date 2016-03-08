if exist "c:\scripts\MSB Update Manager" (
	cd "c:\scripts\MSB Update Manager"
	net stop MSBWindowsUpdateManagement
	"C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe" -u "MSB Windows Update Management.exe"
)

if not exist "c:\scripts\MSB Update Manager" mkdir "c:\scripts\MSB Update Manager"
cd "c:\scripts\MSB Update Manager"
copy "\\itoit22l\b$\Git Workspace\Windows Projects\MSB Windows Update Management\MSB Windows Update Management\bin\Debug\*.*" "c:\scripts\MSB Update Manager"

"C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe" "MSB Windows Update Management.exe"
net start MSBWindowsUpdateManagement

schtasks /Create /RU msb\svckbox /RP D3nv3r0melet2050!! /SC Daily /ST 20:00 /F /RL HIGHEST /TN RefreshWindowsUpdateManagementService /TR "C:\scripts\MSB Update Manager\deploy.bat"

pause