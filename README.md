# QuickRenameExt
## Quick Rename Shell Extension


VB6 made it fairly easy to make shell extensions- ActiveX DLLs that add functionality to the shell. But of course, you're limited to 32bit, and that means 64bit Explorer can't load them. But now with twinBASIC you can compile these as 64bit, as well as have access to features that make it even easier create these extensions.

This first project illustrates how to create a simple Context Menu Handler: an extension that adds options to the right-click menu in Explorer and other apps that show the same context menu. It adds separators and a menu saying "Quick rename", which then has a submenu with several options to automatically format the names of the selected files.

## Using this project

Everything is all ready to go in the .twinproj. Run tB as admin so it can register to HKLM. Then in the toolbar choose 'win64' from the dropdown to compile as 64bit and select build. And that's it-- this project takes advantage of tB's `[RunAfterBuild]` attribute to automatically merge the .reg file.

### Manual registration

There are prebuilt binaries in the Releases section of this repository. You'd need to register with regsvr32 in System32 then double click QuickRenameExt.reg to finish registration.

## How it works 

### Initial setup

Like VB6, you start off by creating a new ActiveX DLL project. All we need in the project is a single class. It's best to use a .twin file here instead of a legacy .cls file because of a handy feature: You can manually specify the Class ID, so you won't need to go looking through the registry to get the GUID for the .reg file we'll need to complete the registration and configuration of our extension at the end.

```vba
(class header)
```

Having access and control over these GUIDs is quite helpful; VB6 hid them from the user. In project settings, you want to make sure 'Use project id as type library id' is checked so you can also control that manually.

Context menu handlers need to implement the COM interfaces `IShellExtInit` and `IContextMenu` at a minimum. While tB supports defining these locally inside the project, the best option is to just add a reference to my 'Windows Development Library for twinBASIC' (WinDevLib). With this you'll have every Windows interface and API you need already defined.

Now you can add `Implements IShellExtInit` and `Implements IContextMenu` then go through and generate the prototypes.

### Getting the selected files list

When someone right clicks and the context menu is created, you get the initialization event in your class in the `IShellExtInit_Init` method. This includes an argument that passes you a pidl for the folded the click occured in and an object representing the files that are selected. They selection needs to be stored in a class-level variable so they're available later. Most examples just look up the file system paths and store those, but in this project I store them as fully qualified pidls, which has the advantage of supporting virtual locations without actual paths.

```vba
(code)
```

### Adding the menu items

The next event that will come through is the construction of the menu itself, where `IContextMenu_QueryContextMenu` is our chance to add our own items. We add a optional separator then it's a straightforward API menu insertion routine.

What's interesting is here we get another major benefit of twinBASIC: If you've seen other VB6 examples of context menu handlers, they employ a technical and complex v-table swap routine to replace the normal class method with an alternate version in a standard module. This is done because of a requirement to return the number of commands added, and since it's a `Sub`, in VB6 you can't directly return a value. In some cases, `Err.Raise` works, but not here. However, tB supports explicitly supplying any return value here via the `Err.ReturnHResult` property, so there's no need to v-table swap it to an alternative.

### Responding to a command 

I'm going to skip over responding to requests for text info; you can see comments in the code that explain that part.\
The main thing of interest is the `IContextMenu_InvokeCommand` method, where we handle a click on one of our menu items. Unlike some other examples, we receive the `CMINVOKECOMMANDINFO` type as a pointer, because it can also be `CMINVOKECOMMANDINFOEX`, and I like to support both. We use tB's new pointer casting support to take a peek at the `.cbSize` member, which tells us which version has been passed:

```vba
(code)
```

The -EX version is supposed to support Unicode, but when testing this project I came across what seems like a bug in Windows; the flag indicating it uses Unicode is set, but the `.lpVerbW` member is 0-- only the original `.lpVerb` contains our command id.

The command id here can be either our numeric id or a String. Since we supplied `.wID` values for the menu items, it should always be the former, but it's good to check to be sure:

```vba
(code}
```

From there we just carry out the desired command. This project uses `IFileOperation` to continue with the high level shell references and methods that let us rename wherever Explorer can, not beholden to requiring a real file system path. It also manages error dialogs automatically.

### Building and registration 

With everything in place, the project can now be built. Like VB6, tB will take care of the standard ActiveX registration, leaving just the need for the shell extension specific registry entries in QuickRenameExt.reg:

```
(reg)
```

The * there means our menu items should appear for all file types. You can specify alternatives here, e.g. changing the * to `exefile` results in the menu only being added for .exe files.

One last helpful new feature, in VB6 you had to find the ProgId in the registry to see what GUID got generated to put in the .reg file, but in tB we manually specified it with the `[ClassId("GUID")]` attribute, which to remind you needs a .twin file. A `[RunAfterBuild]` function will automatically merge the .reg file.

## Conclusion 

That's pretty much it. There's a few other things in the source that should be fairly self-explanatory. Also, the class has a debug logger for the compiled DLL. It's off by default but very helpful for diagnosing problems or just seeing the details of the info the class is receiving.

In future versions, I'll show use of the newer `IExplorerCommand` interface, and there are many other types of shell extensions I want to make demos for.

That's all for now, enjoy!

