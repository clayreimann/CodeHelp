# Version 3.0

All code in this repository is released, with permission from Luthfi, under the
[MIT](https://tldrlegal.com/license/mit-license#summary) license unless otherwise noted.
(For example [CHCore/CHook.cls](CHCore/CHook.cls) or [Plugins/CHTabMDI2/mPublic.bas](Plugins/CHTabMDI2/mPublic.bas)::MakeDWord).
An [installer](CodeHelpInstaller.exe?raw=true) is provided and can be built with [NSIS](http://nsis.sourceforge.net/) 2.46.

### Features
##### CHCoder
  * Autocomplete templates with Shift+Space
  * Use &lt;sel&gt;&lt;/&gt; to delimit the text that should be selected after the expansion

##### CHCodeComplexity
  * The panel displays maximum nesting level,
    [cyclomatic complexity](http://en.wikipedia.org/wiki/Cyclomatic_complexity), and
    [code flow complexity](http://dx.doi.org/10.1109/SCAM.2012.17)
  * The panel should update when switching between components, however Ctrl+M will force an update;
    useful after making some code changes to see if they reduce complexity greatly

##### CHTabMDI
  * When component windows are maximized they are displayed in a rudementary tab interface
  * Ctrl+1-9 will switch between the first 9 windows
  * Ctrl+Enter will maximize the current window

##### CHFullScreen
  * Shift+Enter will remove almost all the window chrome (similiar to zen mode in many other editors)

##### CHMouseWheel
  * Like the solution from Microsoft this plugin enables the mouse wheel in VS6, however this version
    allows the number of lines that are scrolled to be configured

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
