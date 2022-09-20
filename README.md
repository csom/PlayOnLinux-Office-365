# PlayOnLinux-Office-365
Script to install Microsoft Office 365 with PlayOnLinux, this is known to successfully activate Office against Microsoft's servers.


## Compile CX-Wine  
*Hint build Wine on a clean 32-bit virtual machine, I used 32-bit Debian 11 (Bullseye) when building.

- **At first:** download Codeweavers CrossOver sources from: [https://media.codeweavers.com/pub/crossover/source/](https://media.codeweavers.com/pub/crossover/source/)
- Unpack the file corresponding to the version of choseing, in this case: **crossover-sources-21.2.0.tar.gz**
- Before building Wine, add missing ```"dir-of-extracted-source"/sources/wine/include/distversion.h``` file with this content:
```
/* ---------------------------------------------------------------
*   distversion.c
*
* Copyright 2013, CodeWeavers, Inc.
*
* Information from DISTVERSION which needs to find
* its way into the wine tree.
* --------------------------------------------------------------- */

#define WINDEBUG_WHAT_HAPPENED_MESSAGE "This can be caused by a problem in the program or a deficiency in Wine. You may want to check <a href=\"http://www.codeweavers.com/compatibility/\">http://www.codeweavers.com/compatibility/</a> for tips about running this application."

#define WINDEBUG_USER_SUGGESTION_MESSAGE "If this problem is not present under Windows and has not been reported yet, you can save the detailed information to a file using the \"Save As\" button, then <a href=\"http://www.codeweavers.com/support/tickets/enter/\">file a bug report</a> and attach that file to the report."
```
- **Second:** follow the guide at WineHQ how to compile Wine for 32-bit architecture: [https://wiki.winehq.org/Building_Wine](https://wiki.winehq.org/Building_Wine)  
- **Third:** Create an archive with the compiled files:
```
mkdir /tmp/wine-crossover /tmp/wine-crossover/lib /tmp/wine-crossover/share
cp -R /usr/local/bin "dir-of-extracted-source"/sources/wine/include /tmp/wine-crossover
cp -R /usr/local/lib/wine /usr/local/lib/libwine* /tmp/wine-crossover/lib
cp -R /usr/local/share/man /usr/local/share/wine /tmp/wine-crossover/share
cd /tmp/wine-crossover
tar -zcf cx-wine.tar.gz ./*
```

## Use Compiled CX-Wine with POL

- **Fourth:** - Extract the files in: **~/.PlayOnLinux/wine/linux-x86/cxx**


## Install Trial version of CrossOver

- **Fifth:** Install CrossOver 21.2.0
- **Sixth:** Create an empty Windows 7 32-bit bottle inside CrossOver _(This is just to get all the requirements for your system)_
- After this it is possible delete CrossOver and the newley created bottle because this is not needed anymore


## Install Microsoft Office 365
- **Last:** Run the script inside POL: [https://github.com/csom/PlayOnLinux-Office-365/blob/cx-6.0/o365.sh](https://github.com/csom/PlayOnLinux-Office-365/blob/cx-6.0/o365.sh)
