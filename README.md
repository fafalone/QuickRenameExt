# QuickRenameExt
Quick Rename Shell Extension


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
