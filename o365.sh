#!/bin/bash

# Date : (2022-08-23)
# Last revision : (2022-08-23)
# Wine version used : cxx-6.0
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
https://dl.winehq.org/wine/wine-mono/7.0.0/wine-mono-7.0.0-x86.msi

# Choose installer file
POL_SetupWindow_browse "$(eval_gettext 'Please select the setup file to install.')" "$TITLE"
SetupIs="$APP_ANSWER"

# Install Dependencies   
POL_Call POL_Install_msxml6
POL_Call POL_Install_riched20
POL_Call POL_Install_corefonts
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

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\outlook.exe]

[HKEY_CURRENT_USER\Software\Wine\AppDefaults\outlook.exe\DllOverrides]
"activeds"="native,builtin"
"riched20"="native"

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

[HKEY_CURRENT_USER\Software\Wine\Direct2D]
"max_version_factory"=dword:00000000

[HKEY_CURRENT_USER\Software\Wine\DllOverrides]
"*msxml6"=-
"*riched20"=-
"mshtml"="native,builtin"
"msxml6"="native,builtin"
"riched20"="native,builtin"

[HKEY_CURRENT_USER\Software\Wine\Fonts\Replacements]
"@MS UI Gothic"="@Ume UI Gothic"
"Arial"="Bitstream Vera Sans"
"Cambria Math"="Asana Math"
"Gulim"="NanumGothic"
"Lucida Console"="MS Sans Serif"
"MS UI Gothic"="Ume UI Gothic"
"Segoe UI Semilight"="Tahoma"

[HKEY_CURRENT_USER\Software\Wine\MSHTML]

[HKEY_CURRENT_USER\Software\Wine\MSHTML\MainThreadHack]
@=""

[HKEY_CURRENT_USER\Software\Wine\X11 Driver]
"ScreenDepth"="32"' > winesetup.reg

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
	POL_Shortcut "ONENOTE.EXE" "Microsoft OneNote 365" "" "" "Network;InstantMessaging;"
	   
	# NOTE: "Calendar;ContactManagement;"? 
	POL_Shortcut "OUTLOOK.EXE" "Microsoft Outlook 2016" "" "" "Network;Email;"
	
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
	POL_SetupWindow_message "$(eval_gettext '$TITLE has been installed successfully!\n\nThanks!\nBy csoM')" "$TITLE"
fi

# Delete the temporary files and exit
POL_System_TmpDelete
POL_SetupWindow_Close
   

exit 0
