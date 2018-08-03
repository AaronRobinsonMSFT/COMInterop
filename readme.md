# COM Interop

This project is an example on how to manually consume a COM server from C# or a C# server from COM client.

Running COM server with Net client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the COM server (i.e. `ComServer.dll`) using `regsvr32.exe` from an elevated command prompt
    * `regsvr32.exe ComServer.dll`
4) Set the `NetClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the COM server with `regsvr32.exe` passing the `/u` flag (e.g. `regsvr32.exe /u ComServer.dll`).

A project demonstrating Registration Free (RegFree) COM is also included.

Running the RegFree COM server with Net client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
    * The current solution only supports the "F5" experience for the `AnyCPU` and `x86` platforms in RegFree COM.
3) Set the `NetClient_RegFree` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

Running Net server with COM client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the Net server (i.e. `NetServer.dll`) using `regasm.exe` from an elevated command prompt
    * `regasm.exe NetServer.dll /codebase`
    * The `/codebase` flag adds the current path of the assembly to the registry
4) Set the `ComClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the Net server with `regasm.exe` passing the `/u` flag (e.g. `regsvr32.exe /u NetServer.dll`).