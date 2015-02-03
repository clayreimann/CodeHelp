== Version 3.0 ==
All code in this repository is released, with permission from Luthfi, under the MIT license
unless otherwise noted. (For example [CHCore/CHook.cls](CHCore/CHook.cls) or
[Plugins/CHTabMDI2/mPublic.bas](Plugins/CHTabMDI2/mPublic.bas)::MakeDWord).
An installer is provided and will register all the necessary components on a clean system.
It is created with NSIS 2.46.

If you intend to add/extend any of the components the installer expects to find binaries in
the CHCore/bin and CHCore/bin/Plugins directories. If you make any changes to CHGlobalLib you
will need to recompile all of the other dlls (CHCore and all plugins).


== Version 2.2==

To compile the project you need to have WinAPIForVB type library already registered
on your machine. It can be found in the CHCore/Interfaces directory.

*IMPORTANT!!!*
If you decide to compile the add in yourself, please follow these steps:

- Register CHCore/Interfaces/CHLib.tlb
- Register CHCore/Interfaces/WinAPIForVB.tlb
- Build CHGlobal.vbp
- Build CHCore.vbp
- Build all the vbp in plugins sub folder, place the compiled dll of each plugin in "plugins"
  sub folder where the CHCore.dll resides.


Important changes since ver 2.0:

- Moved interface definition from VB.dll to typelib, so the plugin GUID now truly constant.
- Added new plugin (CHTabIdx.dll)
- Added new ShowHelp, ShowPropertyDialog interface
- Added Enable/disable plugin at runtime
- Added features to existing plugin (grouped panels, keyboard shortcut for activating visible tab
  with on-screen hint
- bug fix


Thanks,

Luthfi
