namespace NetClient
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Managed definition of COM interface
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F38720E5-2D64-445E-88FB-1D696F614C78")]
    internal interface IServer
    {
        double ComputePi();
    }

#pragma warning disable IDE1006 // Naming Styles
    /// <summary>
    /// Managed definition of CoClass
    /// </summary>
    [ComImport]
    [CoClass(typeof(ServerClass))]
    [Guid("F38720E5-2D64-445E-88FB-1D696F614C78")]
    internal interface Server : IServer
    {
    }
#pragma warning restore IDE1006 // Naming Styles

    /// <summary>
    /// Managed activation for CoClass
    /// </summary>
    [ComImport]
    [Guid("114383E9-1969-47D2-9AA9-91388C961A19")]
    internal class ServerClass
    {
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Activate COM server
            var server = (Server)new ServerClass();

            var pi = server.ComputePi();
            Console.WriteLine($"\u03C0 = {pi}");
        }
    }
}
