@echo off

REM **********************************************************************
REM ** NAS R2 Setup
REM **
REM ** This command script will install Windows Storage Server 2003 R2
REM ** 
REM ** A command to automatically execute this script is set in the 
REM ** registry during R2 Setup. 
REM ** 
REM ** ?? Syntax:  nasr2setup.cmd [/sr] [/sp]
REM ** ??   /sr = suppress reboot
REM ** ??   /sp = suppress performance optimization changes
REM ** 
REM ** This script will
REM **   o)  Grab Storage Manager settings and put on desktop
REM **   o)  Uninstall NAS 3.0, including Storage Manager+QFE, SAK, and
REM **       the WSS Feature Pack.
REM **   o)  Turn TS On
REM **   o)  Sets wallpaper
REM **   o)  Perf optimzations
REM **   o)  Create a temporary INF file based on the SKU
REM **   o)  Install WSS 2003 R2, through SysOcMgr.  
REM **       Specific SKU components are preinstalled in this step.
REM **   o)  Search
REM **   o)  Boot
REM **
REM **********************************************************************

set DoNotReboot=0
set RebootRequired=0
set NasSku=%1
set DoNotOptimize=0
set CleanInstall=1
set SearchStartType=disabled
set SearchStateRunning=0

if /i "%1" EQU "/sr" set DoNotReboot=1 & set NasSku=%2
if /i "%1" EQU "/sp" set DoNotOptimize=1 & set NasSku=%2

if /i "%2" EQU "/sr" set DoNotReboot=1 & set NasSku=%3
if /i "%2" EQU "/sp" set DoNotOptimize=1 & set NasSku=%3

echo Setting up Windows Storage Server 2003 R2

REM **********************************************************************
REM ** Remove runonce reg key, just in case it is being ran manually
REM **********************************************************************

reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v "NASR2Setup" /f 1>>%windir%\nasr2setup.log 2>&1

REM **********************************************************************
REM ** Turn remote desktop on
REM **********************************************************************

echo Turning on remote desktop

reg add "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v "fDenyTSConnections" /t REG_DWORD /d 0 /f 1>>%windir%\nasr2setup.log 2>&1

REM **********************************************************************
REM ** Detect for clean install
REM **********************************************************************

reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F93FC4CC-E44F-4B57-AFE1-6FA51EC596FC} 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 reg add HKLM\SOFTWARE\Microsoft\WSS /v 2003R2Upgrade /f 1>>%windir%\nasr2setup.log 2>&1

reg query HKLM\SOFTWARE\Microsoft\WSS /v 2003R2Upgrade 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 set CleanInstall=0

REM **********************************************************************
REM ** Grab Storage Manager settings
REM **********************************************************************

REM Check to see if SM is installed

reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F7EF15D0-AB93-48E7-B85A-B80E3C664488} 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% NEQ 0 goto :DoNotOutputSMSettings

echo Writing Storage Manager settings to %userprofile%\Desktop\SM_SaveConfig.txt 

dirquota /show        >> "%userprofile%\Desktop\SM_SaveConfig.txt" 2>&1
filescreening /show   >> "%userprofile%\Desktop\SM_SaveConfig.txt" 2>&1

:DoNotOutputSMSettings

REM **********************************************************************
REM ** To add verbose logging to MsiExec, add the argument: /Lv <logFile>
REM **********************************************************************

REM **********************************************************************
REM ** Uninstall Thunderclap
REM **********************************************************************

echo If present, uninstalling the WSS 2K3 Feature Pack + QFEs

REM WSS Feature Pack (Thunderclap) QFE 838985
MsiExec.exe /X{1E3AA212-B009-485A-9950-DE636AE6C01A} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM WSS Feature Pack (Thunderclap)
MsiExec.exe /X{FE6E4C4D-4C1B-4883-8D17-8FE8D28CAD0A} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM **********************************************************************
REM ** Uninstall MSNFS QFEs
REM **********************************************************************

echo If present, uninstalling the MSNFS QFEs

REM MSNFS QFE: [KB828214]
MsiExec.exe /X{EFBEADC4-822F-4C0A-9241-DCE73171E714} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB828878]                      
MsiExec.exe /X{E3DD998E-2760-4326-8458-390AC29F15CA} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB835152]
MsiExec.exe /X{A7C47891-FF91-4144-9483-60115F75A508} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB838238]
MsiExec.exe /X{C982EDED-0ADA-4F9D-9045-7800C08E9D23} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB887605]
MsiExec.exe /X{BFD83BA7-AC80-4AAF-B9BA-9F0230DA0E3A} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB888618]
MsiExec.exe /X{29774FB6-6E82-4A11-994D-048D596522B8} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB890745]
MsiExec.exe /X{F21AFCAB-1279-4BC9-8F9D-4D079831D5CA} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB891575]
MsiExec.exe /X{BBD6D1D7-B493-498F-B91B-A1AFC5C4062D} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB892159]
MsiExec.exe /X{9DE7625B-F83C-438D-8EA3-8CF8EF8E97D9} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB895398]
MsiExec.exe /X{B3AEDCC1-22B9-4019-828D-EEB02FDCB738} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB894537]
MsiExec.exe /X{12269F79-47A2-4825-AFC6-068B6A86374E} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM MSNFS QFE: [KB897078]
MsiExec.exe /X{12269F79-47A2-4825-AFC6-068B6A86374E} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM **********************************************************************
REM ** Uninstall WQuinn Storage Manager
REM **********************************************************************

echo If present, uninstalling Storage Manager + a set of its QFEs

REM Wquinn QFE: Storage Manager 2.0 for Server Appliances [KB838228]
MsiExec.exe /X{613EC8CB-6C1E-488E-8DFB-1CC4018DC6A3} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM Wquinn QFE: Storage Manager 2.0 for Server Appliances [KB831654]
MsiExec.exe /X{88B0161D-39F2-460C-B3DC-B90F72C85303} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM Wquinn QFE: Manager 2.0 for Server Appliances [KB840141]
MsiExec.exe /X{936A6C89-9E3E-4DAC-8434-92283AC9592F} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM Wquinn QFE: Manager 2.0 for Server Appliances [KB896718]
MsiExec.exe /X{586B5B15-D54E-4089-B8CA-5D28077DEB39} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM Wquinn:  Storage Manager 2.0 for Server Appliances
MsiExec.exe /X{F7EF15D0-AB93-48E7-B85A-B80E3C664488} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM **********************************************************************
REM ** Uninstall SAK
REM **********************************************************************

echo If present, uninstalling the Server Appliance Kit (web page+) + QFEs

REM SAK QFE: [KB832250]
MsiExec.exe /X{4C71616C-51C5-4BB3-8461-EBC7397A8212} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM SAK QFE: [KB842712]
MsiExec.exe /X{95E543A4-C3D0-4C6A-B468-874BAFC4FC27} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM SAK QFE: [KB871116a]
MsiExec.exe /X{1865C1C1-EF9E-4116-AAF2-2DAE7BE773AE} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM Server Appliance Kit 3.0
MsiExec.exe /X{F93FC4CC-E44F-4B57-AFE1-6FA51EC596FC} REBOOT=ReallySuppress /qn
if %errorlevel% EQU 3010 set RebootRequired=1

REM **********************************************************************
REM ** Reboot if one is required for the uninstallation of the above components
REM **********************************************************************

if %RebootRequired% EQU 0 goto :NoUninstallReboot

reg add "HKLM\Software\Microsoft\Windows\CurrentVersion\RunOnce" /v "NASR2Setup" /d "%windir%\system32\nasr2setup.cmd"
shutdown /r /d p:4:1 /t 0 /c "A restart is taking place due to the uninstall of legacy components."
goto :EOF

:NoUninstallReboot

REM **********************************************************************
REM ** Fix up the incomplete uninstall of NFS, but only if R2 NFS is NOT INSTALLED
REM **********************************************************************

reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OC Manager\Subcomponents" /f 1 | findstr /i /c:"nfsserver " 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 goto :R2NfsInstalled

del /f /q %windir%\cluster\nfssh.dll
del /f /q %windir%\cluster\nfsshex.dll 

:R2NfsInstalled

REM **********************************************************************
REM ** Add Performance Optimizations
REM **********************************************************************

if %DoNotOptimize% EQU 1 goto :NoPerfOpt
if %CleanInstall% EQU 0 goto :NoPerfOpt

echo Adding Performance Optimizations

reg add "HKLM\System\CurrentControlSet\Control\FileSystem" /v "NtfsDisableLastAccessUpdate" /t REG_DWORD /d 1 /f 1>>%windir%\nasr2setup.log 2>&1

if %PROCESSOR_ARCHITECTURE% EQU x86 reg add "HKLM\System\CurrentControlSet\Control\Session Manager\Memory Management" /v "PagedPoolSize" /t REG_DWORD /d 192000000 /f 1>>%windir%\nasr2setup.log 2>&1

:NoPerfOpt

REM **********************************************************************
REM ** Set wallpaper
REM **********************************************************************

echo Setting wallpaper

REM wallpaper isn't installed yet, but will install soon

reg add "HKCU\Control Panel\Desktop" /v "Wallpaper" /d "%windir%\web\wallpaper\Windows Storage Server 2003 R2.bmp" /f 1>>%windir%\nasr2setup.log 2>&1
reg add "HKCU\Control Panel\Desktop" /v "WallpaperStyle" /d "2" /f 1>>%windir%\nasr2setup.log 2>&1
reg add "HKCU\Control Panel\Desktop" /v "TileWallpaper" /d "0" /f 1>>%windir%\nasr2setup.log 2>&1

REM Test has shown problems in loading the all users registry hive, as another process is using it.
REM The code below will retry 5 times, and then above; waiting for 5 seconds between each time.

set HiveRetryCount=0
:HiveLoadRetry
reg load HKU\nasr2setup "%allusersprofile%\..\Default User\ntuser.dat" 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 goto :HiveLoaded

echo Cannot load registry hive for default user, will retry up to 5 times
timeout /t 5 /NOBREAK
set /a HiveRetryCount=HiveRetryCount+1
if %HiveRetryCount% LSS 5 goto :HiveLoadRetry

echo Aborting, wallpaper not set for default or new user
goto :HiveDone

:HiveLoaded
reg add "HKU\nasr2setup\Control Panel\Desktop" /v "Wallpaper" /d "%windir%\web\wallpaper\Windows Storage Server 2003 R2.bmp" /f 1>>%windir%\nasr2setup.log 2>&1
reg add "HKU\nasr2setup\Control Panel\Desktop" /v "WallpaperStyle" /d "2" /f 1>>%windir%\nasr2setup.log 2>&1
reg add "HKU\nasr2setup\Control Panel\Desktop" /v "TileWallpaper" /d "0" /f 1>>%windir%\nasr2setup.log 2>&1
reg unload HKU\nasr2setup 1>>%windir%\nasr2setup.log 2>&1
:HiveDone

REM **********************************************************************
REM ** Preserving search configuration
REM **********************************************************************

sc query cisvc | findstr RUNNING 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 set SearchStateRunning=1

sc qc cisvc | findstr START_TYPE | findstr AUTO_START 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 set SearchStartType=auto

sc qc cisvc | findstr START_TYPE | findstr DEMAND_START 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 set SearchStartType=demand

sc qc cisvc | findstr START_TYPE | findstr DISABLED 1>>%windir%\nasr2setup.log 2>&1
if %errorlevel% EQU 0 set SearchStartType=disabled

REM **********************************************************************
REM ** Create a temporary inf file
REM **********************************************************************

REM Not all components will be installed
REM    SRM and DfsFrsUI are only installed on Workgroup+
REM    SanMgmt and Thunderclap only on Standard+
REM This doesn't affect sysoc.inf as the inf won't be there

echo [Components]      > %windir%\inf\nasr2setup.inf
echo FsrNas=on        >> %windir%\inf\nasr2setup.inf
echo PMCSnap=on       >> %windir%\inf\nasr2setup.inf
echo NfsServer=on     >> %windir%\inf\nasr2setup.inf
echo NfsAdminUI=on    >> %windir%\inf\nasr2setup.inf
echo NfsServerAuth=on >> %windir%\inf\nasr2setup.inf
echo NfsClient=on     >> %windir%\inf\nasr2setup.inf
echo MapServer=on     >> %windir%\inf\nasr2setup.inf
echo PortMap=on       >> %windir%\inf\nasr2setup.inf
echo RpcXdr=on        >> %windir%\inf\nasr2setup.inf
echo SRM=on           >> %windir%\inf\nasr2setup.inf
echo DfsFrsUI=on      >> %windir%\inf\nasr2setup.inf
echo DFSR=on          >> %windir%\inf\nasr2setup.inf
echo SanMgmt=on       >> %windir%\inf\nasr2setup.inf 
echo indexsrv_cisvc=on >> %windir%\inf\nasr2setup.inf 

if %CleanInstall% EQU 1  echo sisinst=on        >> %windir%\inf\nasr2setup.inf

REM **********************************************************************
REM ** Run SysOcMgr
REM **********************************************************************

echo Installing the default components for your sku of WSS 2003 R2
echo This may take several minutes, please wait ...

REM Not showing the UI, event though the progress is nice, as we have echo commands

sysocmgr /r /i:%windir%\inf\sysoc.inf /u:%windir%\inf\nasr2setup.inf

if %DoNotReboot% EQU 0 set RebootRequired=1

REM Not suppressing reboot.  If we reboot, then fine, the only remaining step is to reboot.
REM If we suppress reboot we will need to track to figure out how if reboot is required.

REM **********************************************************************
REM ** Set DFSR to Demand
REM **********************************************************************

echo Setting DFSR Service to Demand

sc stop dfsr 1>>%windir%\nasr2setup.log 2>&1
sc config dfsr start= demand 1>>%windir%\nasr2setup.log 2>&1

REM **********************************************************************
REM ** Search configuration
REM **********************************************************************

echo Setting search settings

sc stop cisvc 1>>%windir%\nasr2setup.log 2>&1

reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v CIStateAlwaysAppearsUpToDate /t REG_DWORD /d 1 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v FilterFilesWithUnknownExtensions /t REG_DWORD /d 1 /f 1>>%windir%\nasr2setup.log 2>&1

reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v ScanBackoff /t REG_DWORD /d 0 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxFreshCount /t REG_DWORD /d 100000 /f 1>>%windir%\nasr2setup.log 2>&1
reg delete HKLM\System\CurrentControlSet\Control\ContentIndex /v FilterTrackers /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v FilterTrackers /t REG_MULTI_SZ /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxWordlistIO /t REG_DWORD /d 0xffffffff /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxWordlistIODiskPerf /t REG_DWORD /d 100 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MinimizeWorkingSet /t REG_DWORD /d 0 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxWordlistUserIdle /t REG_DWORD /d 0 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxIndexes /t REG_DWORD /d 50 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MaxWordlists /t REG_DWORD /d 50 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v MinSizeMergeWordlists /t REG_DWORD /d 4096 /f 1>>%windir%\nasr2setup.log 2>&1
reg add HKLM\System\CurrentControlSet\Control\ContentIndex /v GenerateCharacterization /t REG_DWORD /d 0 /f 1>>%windir%\nasr2setup.log 2>&1

if %CleanInstall% EQU 1 goto :CleanInstallSearchConfig

sc config cisvc start= %SearchStartType% 1>>%windir%\nasr2setup.log 2>&1

if %SearchStateRunning% EQU 1 sc start cisvc 1>>%windir%\nasr2setup.log 2>&1

goto :DoneWithSearchConfig

:CleanInstallSearchConfig

sc config cisvc start= auto 1>>%windir%\nasr2setup.log 2>&1

mkdir %systemdrive%\IndexCatalog_System

echo Set objAdminIS = CreateObject("Microsoft.ISAdm")     >  %windir%\nasr2search.vbs
echo objAdminIS.Stop                                      >> %windir%\nasr2search.vbs
echo objAdminIS.RemoveCatalog("System")                   >> %windir%\nasr2search.vbs
echo Set objCatAdmin =  objAdminIS.AddCatalog("System","%systemdrive%\IndexCatalog_System") >> %windir%\nasr2search.vbs
echo objAdminIS.Start                                     >> %windir%\nasr2search.vbs
echo Set objISAdmin = Nothing                             >> %windir%\nasr2search.vbs

%windir%\nasr2search.vbs  1>>%windir%\nasr2setup.log 2>&1

del %windir%\nasr2search.vbs 1>>%windir%\nasr2setup.log 2>&1

:DoneWithSearchConfig

REM **********************************************************************
REM ** Boot if needed
REM **********************************************************************

if %DoNotReboot% EQU 1 goto :EOF
if %RebootRequired% EQU 0 goto :EOF

REM Reducing timeout to 0 seconds to have consistent experience wrt rebooting (see above)
REM Also seen some bugs wrt services being started, it is better this way.

shutdown /r /d p:4:1 /t 0 /c "A restart is taking place due to the installation of WSS 2003 R2."
