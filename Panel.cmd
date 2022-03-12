<!-- : Begin batch script
@setlocal DisableDelayedExpansion
@set uivr=v1
@echo off
:: ### Configuration Options ###

:: change to 1 to enable debug mode (can be used with unattended options)
set _Debug=0

:: change to 0 to turn OFF Windows or Office activation processing via the script
set ActWindows=1
set ActOffice=0

:: change to 0 to turn OFF auto conversion for Office C2R Retail to Volume
set AutoR2V=1

:: change to 0 to revert Windows 10/11 KMS38 to normal KMS
set SkipKMS38=1

:: ### Unattended Options ###

:: change to 1 and set KMS_IP address to activate via external KMS server unattended
set External=0
set KMS_IP=0.0.0.0

:: change to 1 to run Manual activation mode unattended
set uManual=0

:: change to 1 to run AutoRenewal activation mode unattended
set uAutoRenewal=0

:: change to 1 to suppress any output
set Silent=0

:: change to 1 to redirect output to a text file, works only with Silent=1
set Logger=0

:: ### Advanced KMS Options ###

:: change KMS auto renewal schedule, range in minutes: from 15 to 43200
:: example: 10080 = weekly, 1440 = daily, 43200 = monthly
set KMS_RenewalInterval=10080

:: change KMS reattempt schedule for failed activation or unactivated, range in minutes: from 15 to 43200
set KMS_ActivationInterval=120

:: change Hardware Hash for KMS emulator server (only affect Windows 8.1 and 10)
set KMS_HWID=0x3A1C049600B60076

:: change KMS TCP port
set KMS_Port=1688

:: Notice for advanced users on Windows 64-bit (x64 / ARM64):
:: when you bundle KMS_VL_ALL script(s) inside self-extracting program or run it from another command script
:: if the exe pack or the caller script is running as 32-bit (x86) process
:: KMS_VL_ALL script(s) will close then relaunch itself using 64-bit (x64 / ARM64) cmd.exe
:: in that case, be advised not to proceed your pack or caller script depending on KMS_VL_ALL script(s) closure
:: instead, make sure the exe pack or the other caller script are already 64-bit (x64 / ARM64) process

:: ###################################################################
:: # NORMALLY THERE IS NO NEED TO CHANGE ANYTHING BELOW THIS COMMENT #
:: ###################################################################

set KMS_Emulation=1
set Unattend=0
set _uIP=0.0.0.0

set "_Null=1>nul 2>nul"

set "_cmdf=%~f0"
if exist "%SystemRoot%\Sysnative\cmd.exe" (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %*"
exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %*"
exit /b
)

set _args=
set _elev=
set _batf=
set _batp=
set fAUR=
set rAUR=
set _args=%*
if not defined _args goto :NoProgArgs

set _args=%_args:"=%
for %%A in (%_args%) do (
if /i "%%A"=="-elevated" (set _elev=1
) else if /i "%%A"=="/d" (set _Debug=1
) else if /i "%%A"=="/u" (set Unattend=1
) else if /i "%%A"=="/s" (set Silent=1
) else if /i "%%A"=="/l" (set Logger=1
) else if /i "%%A"=="/o" (set ActOffice=1&set ActWindows=0
) else if /i "%%A"=="/w" (set ActOffice=0&set ActWindows=1
) else if /i "%%A"=="/c" (set AutoR2V=0
) else if /i "%%A"=="/x" (set SkipKMS38=0
) else if /i "%%A"=="/e" (set fAUR=0&set External=1&set uManual=0&set uAutoRenewal=0
) else if /i "%%A"=="/m" (set fAUR=0&set External=0&set uAutoRenewal=0
) else if /i "%%A"=="/a" (set fAUR=1&set External=0&set uManual=0
) else if /i "%%A"=="/r" (set rAUR=1
) else (set "KMS_IP=%%A")
)

:NoProgArgs
if %External% EQU 1 (if "%KMS_IP%"=="%_uIP%" (set fAUR=0&set External=0) else (set fAUR=0))
if %uManual% EQU 1 (set fAUR=0&set External=0&set uAutoRenewal=0)
if %uAutoRenewal% EQU 1 (set fAUR=1&set External=0&set uManual=0)
if defined fAUR set Unattend=1
if defined rAUR set Unattend=1
if %Silent% EQU 1 set Unattend=1
set _run=nul
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
set "_err===== ERROR ===="
set "_psc=powershell -nop -c"
set "_buf={$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
set "o_x64=0ca83cdd18845d77e0775f299a111a0591d86883"
set "o_x86=08266ee7d7aac833e04ca037a5435f0438c6b973"
set "o_arm=928aa4e96c49ca3500c714d68ade270e3b4135f5"
set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"&set "_orig=%o_x86%"&set "_wow=0"&set "_bit=32"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"

if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" goto :E_PS

set pwsh32=0
if %xOS%==A64 %_psc% $env:PROCESSOR_ARCHITECTURE 2>nul | find /i "x86" 1>nul && set pwsh32=1
set _dllPath=%SystemRoot%\System32
if %pwsh32% EQU 1 set _dllPath=%SystemRoot%\Sysnative
set _dllNum=1
if %xOS%==x64 set _dllNum=2
if %xOS%==A64 set _dllNum=3
set preparedcolor=0

1>nul 2>nul reg query HKU\S-1-5-19 && (
  goto :Passed
  ) || (
  if defined _elev goto :E_Admin
)

set _PSarg="""%~f0""" %_args% -elevated
set _PSarg=%_PSarg:'=''%

(1>nul 2>nul cscript //NoLogo "%~f0?.wsf" //job:ELAV /File:"%~f0" %_args% -elevated) && (
  exit /b
  ) || (
  call setlocal EnableDelayedExpansion
  1>nul 2>nul %SysPath%\WindowsPowerShell\v1.0\%_psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && (
    exit /b
    ) || (
    goto :E_Admin
  )
)

:Passed
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_Local=%LocalAppData%"
set "_utemp=%TEMP%"
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
set _UNC=0
if "%_work:~0,2%"=="\\" set _UNC=1
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
set "_mO21a=Detected Office 2021 C2R Retail is activated"
set "_mO19a=Detected Office 2019 C2R Retail is activated"
set "_mO16a=Detected Office 2016 C2R Retail is activated"
set "_mO15a=Detected Office 2013 C2R Retail is activated"
set "_mO21c=Detected Office 2021 C2R Retail could not be converted to Volume"
set "_mO19c=Detected Office 2019 C2R Retail could not be converted to Volume"
set "_mO16c=Detected Office 2016 C2R Retail could not be converted to Volume"
set "_mO15c=Detected Office 2013 C2R Retail could not be converted to Volume"
set "_mO14c=Detected Office 2010 C2R Retail is not supported by KMS_VL_ALL"
set "_mO14m=Detected Office 2010 MSI Retail is not supported by KMS_VL_ALL"
set "_mO15m=Detected Office 2013 MSI Retail is not supported by KMS_VL_ALL"
set "_mO16m=Detected Office 2016 MSI Retail is not supported by KMS_VL_ALL"
set "_mOuwp=Detected Office 365/2016 UWP is not supported by KMS_VL_ALL"
set DO16Ids=ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,Excel,Outlook,PowerPoint,Publisher,Word
set LV16Ids=Mondo,ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,OneNote,Excel,Outlook,PowerPoint,Publisher,Word
set LR16Ids=%LV16Ids%,Professional,HomeBusiness,HomeStudent,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set "ESUEditions=Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
set "ESUEditions=ServerDatacenter,ServerDatacenterCore,ServerDatacenterV,ServerDatacenterVCore,ServerStandard,ServerStandardCore,ServerStandardV,ServerStandardVCore,ServerEnterprise,ServerEnterpriseCore,ServerEnterpriseV,ServerEnterpriseVCore"
)
setlocal EnableDelayedExpansion

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  set "_Pause=pause >nul"
  if %Unattend% EQU 1 set "_Pause="
  if %Silent% EQU 0 (call :Begin) else (call :Begin >!_run! 2>&1)
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set "_Pause="
  copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_log=!_dsk!\%~n0")
  if exist "!_log!_Debug.log" (
  call set "_suf="
  for /f "tokens=2 delims==." %%# in ('wmic os get localdatetime /value') do set "_date=%%#"
  set "_suf=_!_date:~8,6!"
  )
  if %Silent% EQU 0 (
  echo.
  echo Running in Debug Mode...
  if not defined _args (echo The window will be closed when finished) else (echo please wait...)
  echo.
  echo writing debug log to:
  echo "!_log!_Debug!_suf!.log"
  )
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug!_suf!.log"&del "!_log!_tmp.log"
)
@color 07
@title %ComSpec%
@echo off
@exit /b

:Begin
if %_Debug% EQU 1 (
if defined _args echo %_args%
echo "!_batf!"
)
if exist "%PUBLIC%\ReadMeAIO.html" del /f /q "%PUBLIC%\ReadMeAIO.html"
if exist "%_temp%\'" del /f /q "%_temp%\'"
if exist "%_temp%\`.txt" del /f /q "%_temp%\`.txt"
set _verb=0
set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _Hook="%SysPath%\SppExtComObjHook.dll"
set w7inf=%SystemRoot%\Migration\WTR\KMS_VL_ALL.inf
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"
set "_TaskOs=\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTaskLogon"
set "line1============================================================="
set "line2=************************************************************"
set "line3=____________________________________________________________"
set "line4=__________________________________________________"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% EQU 1060 set OsppHook=0

set ESU_KMS=0
if %winbuild% LSS 9200 for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\channels') do (
  if exist "%SysPath%\spp\tokens\channels\%%A\*VL-BYPASS*.xrm-ms" set ESU_KMS=1
)
if %ESU_KMS% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "addon=and LicenseDependsOn is not NULL") else (set "adoff="&set "addon=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (%ESUEditions%) do (
  if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
)
if %ESU_EDT% EQU 1 set SSppHook=1
set ESU_ADD=0

if %winbuild% GEQ 9200 (
  set OSType=Win8
  set SppVer=SppExtComObj.exe
) else if %winbuild% GEQ 7600 (
  set OSType=Win7
  set SppVer=sppsvc.exe
) else (
  goto :UnsupportedVersion
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %_Nul3% && (
reg delete "%IFEO%\sppsvc.exe" /f %_Nul3%
call :StopService sppsvc
)
set _uRI=%KMS_RenewalInterval%
set _uAI=%KMS_ActivationInterval%
set _dDbg=No
if %ActWindows% EQU 0 if %ActOffice% EQU 0 set ActWindows=1
if %_Debug% EQU 1 if not defined fAUR set fAUR=0&set External=0
if %Unattend% EQU 1 if not defined fAUR set fAUR=0&set External=0
if not defined fAUR if not defined rAUR goto :MainMenu
if defined rAUR (set _verb=1&cls&call :RemoveHook&goto :cCache)
set Unattend=1
set _ReAR=0
set _AUR=0
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || (
  reg query "%IFEO%\%SppVer%" /v VerifierFlags %_Nul3% && set _AUR=1
  if %SSppHook% EQU 0 reg query "%IFEO%\osppsvc.exe" /v VerifierFlags %_Nul3% && set _AUR=1
)
if %fAUR% EQU 1 (set _ReAR=1&if %_AUR% EQU 0 (set _AUR=1&set _verb=1&set _rtr=DoActivate&cls&goto :InstallHook) else (set _verb=0&set _rtr=DoActivate&cls&goto :InstallHook))
if %External% EQU 0 (set _AUR=0&cls&goto :DoActivate)
cls&goto :DoActivate


:MainMenu
@mode con cols=90 lines=50
@echo off
@color a
@echo ----------------------------------------------------------------------------------
@echo	 SSSSS    .sSSSSSSSSs.   .sSSSsSS SSsSSSSS .sSSSSs.    .sSSS SSSSS .sSSSSs.    
@echo	 SSSSS .sSSSSSSSSSSSSSs. SSSSS  SSS  SSSSS SSSSSSSSSs. SSSSS SSSSS SSSSSSSSSs. 
@echo	 S SSS SSSSS S SSS SSSSS S SSS   S   SSSSS S SSS SSSS' S SSS SSSSS S SSS SSSSS 
@echo	 S  SS SSSSS S  SS SSSSS S  SS       SSSSS S  SS       S  SS SSSSS S  SS SSSS' 
@echo	 S..SS `:S:' S..SS `:S:' S..SS       SSSSS S..SSsss    S..SSsSSSSS S..SSsSSSa. 
@echo	 S:::S       S:::S       S:::S       SSSSS S:::SSSS    S:::S SSSSS S:::S SSSSS 
@echo	 S;;;S       S;;;S       S;;;S       SSSSS S;;;S       S;;;S SSSSS S;;;S SSSSS 
@echo	 S:::S       S:::S       S:::S       SSSSS S:::SSSS    S:::S SSSSS S:::S SSSSS 
@echo	 SSSSS       SSSSS       SSSSS       SSSSS SSSSSsSSS:' SSSSS SSSSS SSSSS SSSSS 
                                                                              
@echo ----------------------------------------------------------------------------------
echo   Programmed by ITMEHR CyberSecurity team
set m=https://github.com/itmehr
set n=https://instagram.com/itmehr
echo   Follow us on  [31m%m%[0m
echo  [32m Follow us on  [31m%n%[0m
echo [32m-------------------------------------------------
echo	 Select once:
echo  [0m	1] Enable Wi-Fi Network Adapter 
echo  	2] Disable Wi-Fi Network Adapter [0m
echo .
echo 	3] Erase History Leaked from Incognito Browser Mode 
echo .
echo 	4] Start Hotspot Network
echo 	5] Stop Hotspot Network
echo .
echo 	6] Open Network Adapters
echo .
echo 	7] Show Wi-Fi Signal Streng
echo .
choice /c 12345670 /n /m "  	> Choose a menu option, or press 0 to Exit: "


set _el=%errorlevel%
if %_el%==7 (cls&goto :Signal)
if %_el%==6 (cls& ncpa.cpl & goto :MainMenu)
if %_el%==5 (cls&goto :unspot)
if %_el%==4 (cls&goto :spot)
if %_el%==3 (cls&goto :Flush)
if %_el%==2 (cls&goto :DeActivate)
if %_el%==1 (cls&goto :Activate)
goto:eof


:Activate
@netsh interface set interface name="Wi-Fi" admin=enabled>nul
echo   -Wi-Fi has been Enabled
timeout /t 5 & goto:MainMenu


:DeActivate
@netsh interface set interface name="Wi-Fi" admin=disabled>nul
echo   -Wi-Fi has been Disabled
timeout /t 5 & goto:MainMenu

:Flush
@ipconfig/flushdns>nul
echo   -Your Leaked History From Incongito Browser Mode Has Been Erased!
timeout /t 5 & goto:MainMenu


:unspot
@netsh wlan stop hostednetwork >nul
echo   -Hotspot "%id%":"%pass%" Has Been Stopped!
timeout /t 5 & goto:MainMenu



:spot
set /p id="> Enter SSID:"
if %id% EQU nul goto:spot
set /p pass="> Enter PASSWORD:"
if %pass% EQU nul goto:spot
@netsh wlan set hostednetwork ssid="%id%" key="%pass%" >nul
@netsh wlan start hostednetwork >nul
echo   -Hotspot "%id%":"%pass%" Has Been Started!
timeout /t 5 & goto:MainMenu


:Signal
@ECHO "> Press Ctrl + C to Cancel;"
:one
netsh wlan show interfaces | findstr "Signal"
timeout 1 >nul | goto :two

:two
goto :one


:E_Admin
echo %_err%
echo This script requires administrator privileges.
echo To do so, right-click on this script and select 'Run as administrator'
echo.
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
pause >nul
goto :eof



:E_PS
echo %_err%
echo Windows PowerShell is required for this script to work.
echo.
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
pause >nul
goto :eof

----- Begin wsf script --->
<package>
   <job id="ELAV">
       <script language="VBScript">
           Set strArg=WScript.Arguments.Named
           If Not strArg.Exists("File") Then
               Wscript.Echo "Switch /File:<File> is missing."
               WScript.Quit 1
           End If
           Set strRdlproc = CreateObject("WScript.Shell").Exec("rundll32 kernel32,Sleep")
           With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & strRdlproc.ProcessId & "'")
               With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & .ParentProcessId & "'")
                   If InStr (.CommandLine, WScript.ScriptName) <> 0 Then
                       strLine = Mid(.CommandLine, InStr(.CommandLine , "/File:") + Len(strArg("File")) + 8)
                   End If
               End With
               .Terminate
           End With
          CreateObject("Shell.Application").ShellExecute "cmd.exe", "/c " & chr(34) & chr(34) & strArg("File") & chr(34) & strLine & chr(34), "", "runas", 1
       </script>
   </job>
</package>


:eof