# Solid Edge Top Level Assembly Opener

This is a Windows Explorer shell extension which will give you a new right-click menu on Windows directories. When selecting Solid Edge > Open Top-Level Assemblies, the shell extension will search the directory for the existence of files with the .asm extension. If any are found, it will search for a running Solid Edge application to attach to or it will launch a new one, then open all of the top-level assembly files found in the directory then set the Solid Edge application to default foreground settings.

To install, build or download the SolidEdgeTopLevelAsmOpener.dll file and place in a directory of your choice. Use regasm.exe to register the .dll and you should see the menu available when right clicking a file in Explorer.
