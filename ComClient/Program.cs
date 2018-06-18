namespace ComClient
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
        static IServer[] Servers = new IServer[8];

        void AccessServers()
        {
            foreach (var s in Servers)
            {
                var pi = s.ComputePi();
                Console.WriteLine($"PI: {pi}");
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            // Activate COM server
            for (int i = 0; i < Servers.Length; ++i)
            {
                Servers[i] = (IServer)new ServerClass();
            }

            new Program().AccessServers();

            // Null out all references or else the GC won't collect RCWs
            Servers = null;

            GC.Collect();
            GC.Collect();

            // In mixed mode debugging, the final release for the COM objects
            // can be observed with WaitForPendingFinalizers() on the stack.
            GC.WaitForPendingFinalizers();
        }
    }
}
