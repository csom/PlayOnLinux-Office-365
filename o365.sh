#!/bin/bash

# Date : (2022-08-23)
# Last revision : (2022-08-29)
# Wine version used : cx-6.0
# Distribution used to test : Ubuntu 22.04 LTS
# Author : csoM
# PlayOnLinux : 4.3.4
# Script licence : csoM
# Program licence : Test

# Initialization!
[ "$PLAYONLINUX" = "" ] && exit 0
source "$PLAYONLINUX/lib/sources"
   
TITLE="Microsoft Office 365"
PREFIX="Office365"
WINEVERSION="cxx"
ARCH="x86"
   
POL_GetSetupImages "https://i.imgur.com/licFVuF.png" "https://i.imgur.com/ff6PkEZ.png" "$TITLE"
   
POL_SetupWindow_Init
POL_SetupWindow_SetID 3066

POL_SetupWindow_message "$(eval_gettext 'Please make sure to have CodeWeavers Wine version 21.2.0 installed in location ".PlayOnLinux/wine/linux-x86/cxx" before you continue with your installation.\n\nThanks!\nBy csoM')" "$TITLE"
   
POL_SetupWindow_presentation "$TITLE" "Microsoft" "http://www.microsoft.com" "csoM" "$PREFIX"
   
POL_Debug_Init
POL_System_TmpCreate "$PREFIX" 
# ---------------------------------------------------------------------------------------------------------
# Perform some validations!
POL_RequiredVersion 4.3.4 || POL_Debug_Fatal "$TITLE won't work with $APPLICATION_TITLE $VERSION!nPlease update!"
   
#Linux
if [ "$POL_OS" = "Linux" ]; then
    wbinfo -V || POL_Debug_Fatal "Please install winbind before installing $TITLE!"
else
    POL_Debug_Fatal "$(eval_gettext "Only Linux OS is supported! Sorry!")";
    POL_SetupWindow_Close
    exit 1
fi

 
#Validation of 32Bits
if [ ! "$(file $SetupIs | grep 'x86-64')" = "" ]; then
    POL_Debug_Fatal "$(eval_gettext "The 64bits version is not compatible! Sorry!")";
fi
   
# ---------------------------------------------------------------------------------------------------------
# Prepare resources for installation!

# Prepare the Wine Prefix
POL_System_SetArch "$ARCH"  
POL_Wine_SelectPrefix "$PREFIX"
POL_Wine_PrefixCreate "$WINEVERSION"

# Download MSXML 6.0 SP2
DOWNLOAD_URL="https://catalog.s.download.windowsupdate.com/msdownload/update/software/crup/2009/11/msxml6-kb973686-enu-x86_e139664a78bc2806cf0c5bcf0bedec7ea073c3b1.exe"
cd "$POL_System_TmpDir"
POL_Download "$DOWNLOAD_URL"
DOWNLOAD_FILE="$(basename "$DOWNLOAD_URL")"
SetupSP2Is="$POL_System_TmpDir/$DOWNLOAD_FILE"


# Download Wine Mono 7.0.0
DOWNLOAD_URL="https://dl.winehq.org/wine/wine-mono/7.0.0/wine-mono-7.0.0-x86.msi"
cd "$POL_System_TmpDir"
POL_Download "$DOWNLOAD_URL"
DOWNLOAD_FILE="$(basename "$DOWNLOAD_URL")"
SetupMonoIs="$POL_System_TmpDir/$DOWNLOAD_FILE"

# Download Asana Font
DOWNLOAD_URL="https://ctan.org/tex-archive/fonts/Asana-Math/ASANA.TTC"
cd "$POL_System_TmpDir"
POL_Download "$DOWNLOAD_URL"
DOWNLOAD_FILE="$(basename "$DOWNLOAD_URL")"
ASANAFONT="$POL_System_TmpDir/$DOWNLOAD_FILE"

# Choose installer file
POL_SetupWindow_browse "$(eval_gettext 'Please select the setup file to install.')" "$TITLE"
SetupIs="$APP_ANSWER"

# Install Dependencies   
POL_Call POL_Install_msxml6
POL_Call POL_Install_riched20
POL_Call POL_Install_corefonts
cp "$ASANAFONT" "$WINEPREFIX/drive_c/windows/Fonts/"
POL_AutoWine "$SetupMonoIs"

# Change to Windows 2000 to allow MSCXML 6.0 SP2 to be installed
cd "$POL_System_TmpDir"
echo -e 'REGEDIT4

[HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion]
"CSDVersion"="Service Pack 4"
"CurrentBuild"="2195"
"CurrentBuildNumber"="2195"
"CurrentMajorVersionNumber"=dword:00000005
"CurrentMinorVersionNumber"=dword:00000000
"CurrentType"="Uniprocessor Free"
"CurrentVersion"="5.0"
"DigitalProductId"=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00
"InstallDate"=dword:4be5019a
"ProductId"="12345-oem-0000001-54321"
"ProductName"="Microsoft Windows 2000"
"RegisteredOrganization"=""
"RegisteredOwner"=""
"SystemRoot"="C:\\windows"' > win2k.reg

POL_Wine regedit win2k.reg

# Install SP2
POL_Wine "$SetupSP2Is"


# Change back to Windows 7
cd "$POL_System_TmpDir"
echo -e 'REGEDIT4

[HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion]
"CSDVersion"="Service Pack 1"
"CurrentBuild"="7601"
"CurrentBuildNumber"="7601"
"CurrentMajorVersionNumber"=dword:00000006
"CurrentMinorVersionNumber"=dword:00000001
"CurrentType"="Uniprocessor Free"
"CurrentVersion"="6.1"
"DigitalProductId"=hex:00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,\
  00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00
"InstallDate"=dword:4be5019a
"ProductId"="12345-oem-0000001-54321"
"ProductName"="Microsoft Windows 7"
"RegisteredOrganization"=""
"RegisteredOwner"=""
"SystemRoot"="C:\\windows"' > win7.reg

POL_Wine regedit win7.reg


# Prepare the Registry with necessary overrides 
cd "$POL_System_TmpDir"
echo -e 'REGEDIT4

[HKEY_CURRENT_USER\Software\Wine]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\MCC-Win64-Shipping.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\MCC-Win64-Shipping.exe\DllOverrides]
"concrt140"="native, builtin"

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\outlook.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\outlook.exe\DllOverrides]
"activeds"="native,builtin"
"riched20"="native"

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\vc_redist.x64.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\vc_redist.x64.exe\DllOverrides]
"msxml2"="builtin, native"
"msxml3"="builtin, native"

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\winemenubuilder.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\winemenubuilder.exe\\Explorer]
"Desktop"="root"

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\winewrapper.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\winewrapper.exe\DllOverrides]
"crypt32"="builtin"
"rsabase"="builtin"
"rsaenh"="builtin"

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\winewrapper.exe\\Explorer]
"Desktop"="root"

[HKEY_CURRENT_USER\Software\Wine\Debug]
"RelayExclude"="ntdll.LdrFindResource_U;ntdll.NtCreateSection;ntdll.NtMapViewOfSection;ntdll.NtQueryAttributesFile;ntdll.NtClose;ntdll.NtOpenFile;ole32.CoUninitialize;ntdll.RtlEnterCriticalSection;ntdll.RtlLeaveCriticalSection;kernel32.48;kernel32.49;kernel32.94;kernel32.95;kernel32.96;kernel32.97;kernel32.98;kernel32.TlsGetValue;kernel32.TlsSetValue;kernel32.FlsGetValue;kernel32.FlsSetValue;kernel32.SetLastError"
"RelayFromExclude"=""

[HKEY_CURRENT_USER\Software\Wine\Direct2D]
"max_version_factory"=dword:00000000

[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"*msxml6"=-
"*riched20"=-
"*autorun.exe"="native,builtin"
"*ctfmon.exe"="builtin"
"*ddhelp.exe"="builtin"
"*docbox.api"=""
"*findfast.exe"="builtin"
"*ieinfo5.ocx"="builtin"
"*maildoff.exe"="builtin"
"*mdm.exe"="builtin"
"*mosearch.exe"="builtin"
"*msiexec.exe"="builtin"
"*pstores.exe"="builtin"
"*user.exe"="native,builtin"
"amstream"="native, builtin"
"atl"="native, builtin"
"concrt140"="native"
"crypt32"="native, builtin"
"d3dxof"="native, builtin"
"dciman32"="native"
"devenum"="native, builtin"
"dplay"="native, builtin"
"dplaysvr.exe"="native, builtin"
"dplayx"="native, builtin"
"dpnaddr"="native, builtin"
"dpnet"="native, builtin"
"dpnhpast"="native, builtin"
"dpnhupnp"="native, builtin"
"dpnlobby"="native, builtin"
"dpnsvr.exe"="native, builtin"
"dpnwsock"="native, builtin"
"dxdiagn"="native, builtin"
"hhctrl.ocx"="native, builtin"
"hlink"="native, builtin"
"iernonce"="native, builtin"
"itss"="native, builtin"
"jscript"="native, builtin"
"mlang"="native, builtin"
"mshtml"="native, builtin"
"msi"="builtin"
"msvcirt"="native, builtin"
"msvcp140"="native,builtin"
"msvcrt40"="native, builtin"
"msvcrtd"="native, builtin"
"msxml6"="native, builtin"
"odbc32"="native, builtin"
"odbccp32"="native, builtin"
"ole32"="builtin"
"oleaut32"="builtin"
"olepro32"="builtin"
"quartz"="native, builtin"
"riched20"="native,builtin"
"riched32"="native, builtin"
"rpcrt4"="builtin"
"rsabase"="native, builtin"
"secur32"="native, builtin"
"shdoclc"="native, builtin"
"shdocvw"="native, builtin"
"softpub"="native, builtin"
"urlmon"="native, builtin"
"wininet"="builtin"
"wintrust"="native, builtin"
"wscript.exe"="native, builtin"

[HKEY_CURRENT_USER\Software\Wine\\EnableOLEQuitFix]
@=""

[HKEY_CURRENT_USER\Software\Wine\\EnableUIAutomationCore]
@=""

[HKEY_CURRENT_USER\Software\Wine\Fonts\Replacements]
"Arial"="FreeSans"
"Cambria Math"="Asana Math"
"Lucida Console"="FreeSerif"
"Segoe UI Semilight"="Tahoma"

[HKEY_CURRENT_USER\Software\Wine\Mac Driver]
"OpenGLSurfaceMode"="behind"

[HKEY_CURRENT_USER\Software\Wine\MSHTML]

[HKEY_CURRENT_USER\Software\Wine\MSHTML\MainThreadHack]
@=""

[HKEY_CURRENT_USER\Software\Wine\X11 Driver]
"ScreenDepth"="32"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration]
"CDNBaseUrl"="http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"' > winesetup.reg

POL_Wine regedit winesetup.reg
   
# ---------------------------------------------------------------------------------------------------------
# Install!
if [[ "$SetupIs" = *"exe"* ]]
then
	POL_Wine "$SetupIs"

	# ---------------------------------------------------------------------------------------------------------
	# Create shortcuts, entries to extensions and finalize!
	   
	# NOTE: Create shortcuts! 
	POL_Shortcut "WINWORD.EXE" "Microsoft Word 365" "" "" "Office;WordProcessor;"
	POL_Shortcut "EXCEL.EXE" "Microsoft Excel 365" "" "" "Office;Spreadsheet;"
	POL_Shortcut "POWERPNT.EXE" "Microsoft Powerpoint 365" "" "" "Office;Presentation;"
	   
	# NOTE: No category for collaborative work? 
	POL_Shortcut "ONENOTE.EXE" "Microsoft OneNote 2016" "" "" "Network;InstantMessaging;"
	   
	# NOTE: "Calendar;ContactManagement;"? 
	POL_Shortcut "OUTLOOK.EXE" "Microsoft Outlook 365" "" "" "Network;Email;"
	
	# NOTE: "Storage;"? 
	#POL_Shortcut "OneDrive.EXE" "Microsoft OneDrive 365" "" "" "Network;Storage;"
	   
	# NOTE: Add an entry to PlayOnLinux's extension file. If the entry already
	# exists, it will replace it! By Questor
	# [Ref.: https://github.com/PlayOnLinux/POL-POM-4/blob/master/lib/playonlinux.lib]
	#POL_Extension_Write doc "Microsoft Word 365"
	#POL_Extension_Write docx "Microsoft Word 365"
	#POL_Extension_Write xls "Microsoft Excel 365"
	#POL_Extension_Write xlsx "Microsoft Excel 365"
	#POL_Extension_Write ppt "Microsoft Powerpoint 365"
	#POL_Extension_Write pptx "Microsoft Powerpoint 365"
	
	# Change to manual account creation in Outlook
	cd "$POL_System_TmpDir"
	echo -e 'REGEDIT4

[HKEY_USERS\S-1-5-21-0-0-0-1000\Software\Microsoft\Office\16.0\PowerPoint\Options]
"DisableBootToOfficeStart"=dword:00000001

[HKEY_USERS\S-1-5-21-0-0-0-1000\Software\Microsoft\Office\16.0\Word\Options]
"DisableBootToOfficeStart"=dword:00000001

[HKEY_USERS\S-1-5-21-0-0-0-1000\Software\Microsoft\Office\16.0\\Excel\Options]
"DisableBootToOfficeStart"=dword:00000001

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Setup]
"DisableOffice365SimplifiedAccountCreation"=dword:00000001' > postreg.reg

	POL_Wine regedit postreg.reg

	
	POL_SetupWindow_message "$(eval_gettext '$TITLE has been installed successfully!\n\nThanks!\nBy csoM')" "$TITLE"
	
fi

# Delete the temporary files and exit
POL_System_TmpDelete
POL_SetupWindow_Close
   

exit 0
