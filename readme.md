# COM Interop

This project is an example on how to manually consume a C++ COM server from C#.

Running example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the COM server (i.e. `ComServer.dll`) using `regsvr32.exe` from an elevated command prompt
    * `regsvr32.exe ComServer.dll`
4) Set the `ComClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the COM server with `regsvr32.exe` passing the `/u` flag (e.g. `regsvr32.exe /u ComServer.dll`).
