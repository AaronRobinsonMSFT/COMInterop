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

Running Net server with COM client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Register the Net server (i.e. `NetServer.dll`) using `regasm.exe` from an elevated command prompt
    * `regasm.exe NetServer.dll /codebase`
    * The `/codebase` flag adds the current path of the assembly to the registry
4) Set the `ComClient` project as the StartUp project
5) Press "F5" from within Visual Studio to debug

When done with the project, remember to unregister the Net server with `regasm.exe` passing the `/u` flag (e.g. `regsvr32.exe /u NetServer.dll`).

Projects demonstrating Registration Free (RegFree) COM are also included.

Running the RegFree COM server with Net client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
    * The current solution only supports the "F5" experience for the `AnyCPU` and `x86` platforms in RegFree COM.
3) Set the `NetClient_RegFree` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

Running the RegFree Net server with COM client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x86`)
3) Set the `ComClient_RegFree` project as the StartUp project
4) Press "F5" from within Visual Studio to debug

Running the RegFree Net Core Server with COM client example:

1) Load ComInterop.sln in Visual Studio
2) Build desired solution configuration (e.g. `Debug|x64`)
    * Only x64 configurations are supported
3) Set the `ComClient_RegFree_Core` project as the StartUp project
4) Open project properties and select the 'Debugging' page
5) Change the working directory from `$(ProjectDir)` to `$(TargetDir)`
6) Set the `CORE_ROOT` environment variable to the absolute path of the desired .NET Core framework
    * i.e. `CORE_ROOT=C:\Program Files\dotnet\shared\Microsoft.NETCore.App\2.1.2`
7) Apply changes and close project properties
8) Press "F5" from within Visual Studio to debug

## References

[RegFree COM Walkthrough](https://msdn.microsoft.com/en-us/library/ms973913.aspx)

[RegFree COM with .NET Framework](https://docs.microsoft.com/en-us/dotnet/framework/interop/configure-net-framework-based-com-components-for-reg)

[Hosting .NET Core on Unix](https://docs.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting#about-hosting-net-core-on-unix)

[COM Activation in .NET Core proposal](https://github.com/dotnet/core-setup/pull/4476)
