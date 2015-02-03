# Version 3.0

All code in this repository is released, with permission from Luthfi, under the MIT license
unless otherwise noted. (For example [CHCore/CHook.cls](CHCore/CHook.cls) or
[Plugins/CHTabMDI2/mPublic.bas](Plugins/CHTabMDI2/mPublic.bas)::MakeDWord).
An [installer](CodeHelpInstaller.exe?raw=true) is provided and can be built with [NSIS](http://nsis.sourceforge.net/) 2.46.

### Extending CodeHelp
If you intend to add/extend any of the components the installer expects to find binaries in
the `CHCore/bin` and `CHCore/bin/Plugins` directories. You should copy the example project in `Plugins/BlankTemplate` to a
new subdirectory and add it to CodeHelp.vbg. 

To compile the project you need to have the `WinAPIForVB` and `chlib` type libraries registered on your machine. They are 
located in the CHCore/Interfaces directory. Then you will first need to compile `CHGlobalLib` because `CHCore` and all the
plugins reference `CHGlobalLib`. `CHGlobalLib` and `CHCore` go in `CHCore/bin`, all the plugins go in `CHCore/bin/Plugins`
and must be registered as COM servers.

*Note:* If you make any changes to `CHGlobalLib` you will need to recompile all of the other dlls (CHCore and all plugins).


##### Contributors
Thanks to Luthfi for developing this initally and allowing it to be revived.
