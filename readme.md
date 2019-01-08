# COM Interop

This project is an example on how to manually consume a COM server from C# or a C# server from COM client. It also contains projects for less common scenarios involving .NET and COM.

Running COM server with Net client example:

1) Load `ComInterop.sln` in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the COM server (i.e. `ComServer.dll`) using `regsvr32.exe` from an elevated command prompt
    * `regsvr32.exe ComServer.dll`
4) Set the `NetClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the COM server with `regsvr32.exe` passing the `/u` flag (e.g. `regsvr32.exe /u ComServer.dll`).

Running Net server with COM client example:

1) Load `ComInterop.sln` in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the Net server (i.e. `NetServer.dll`) using `regasm.exe` from an elevated command prompt
    * `regasm.exe NetServer.dll /codebase`
    * The `/codebase` flag adds the current path of the assembly to the registry
4) Set the `ComClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the Net server with `regasm.exe` passing the `/u` flag (e.g. `regsvr32.exe /u NetServer.dll`).

Projects demonstrating Registration Free (RegFree) COM are also included.

Running the RegFree COM server with Net client example:

1) Load `ComInterop.sln` in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
    * The current solution only supports the "F5" experience for the `AnyCPU` and `x86` platforms in RegFree COM.
3) Set the `NetClient_RegFree` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

Running the RegFree Net server with COM client example:

1) Load `ComInterop.sln` in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Set the `ComClient_RegFree` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

Running the Out-of-proc demo:

1) Load the `ComInterop.sln` in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Set the `OutOfProcDemo` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

**Note** The Out-of-proc demo launches a child process from the main process.

## References

[RegFree COM Walkthrough](https://msdn.microsoft.com/library/ms973913.aspx)

[RegFree COM with .NET Framework](https://docs.microsoft.com/dotnet/framework/interop/configure-net-framework-based-com-components-for-reg)

[Running Object Table](https://docs.microsoft.com/windows/desktop/api/objidl/nn-objidl-irunningobjecttable)

[Type Libraries](https://msdn.microsoft.com/library/windows/desktop/ms221060.aspx)
