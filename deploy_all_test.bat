rem @echo off
cls
REM ______________________________________________________________

echo setting up tasks > log.txt

REM Set your server list here
set SERVERLIST=govsqltst logossqltest msbsqltst sql2012dev sdesql10tst logosrpttest sqlrpt msbrpttst logostest isupptst trim7train trim7test gis102agstst spapp01tst spsql01tst spsql02tst spweb01tst spweb02tst

:GetBestServer
set command=\\dsjkb\desoft$\MSBWindowsUpdateManagementSvc\deploy.bat
for %%s in (%SERVERLIST%) do (
	echo %%s >> log.txt
	schtasks /S \\%%s /Create /RU msb\svckbox /RP D3nv3r0melet2050!! /SC ONCE /ST 09:50 /F /RL HIGHEST /TN DeployWindowsUpdateManagementService /TR %command%
	rem at \\%%s 09:25 %command% >> log.txt
)

pause