Version 2.2

To compile the project you need to have WinAPIForVB type library already registered
on your machine. It can be found in the CHCore/Interfaces directory.

IMPORTANT!!!
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
