using IServerLib;

namespace OutOfProcDemo
{
    using System;
    using System.Diagnostics;
    using System.Reflection;
    using System.Runtime.InteropServices;

    class Program
    {
        private static int processId;

        static void Main(string[] args)
        {
            // Record the current process ID
            using (var p = Process.GetCurrentProcess())
            {
                processId = p.Id;
            }

            const string remoteObjectName = "OutOfProcDemoObject";

            // Acquire the Running Object Table instance
            using (var rot = new RunningObjectTable())
            {
                if (args.Length == 0 || args[0].Equals("server"))
                {
                    Log($"Server running...");

                    // Register object instance in Running Object Table
                    var server = new NetServer.Server();
                    using (rot.Register(remoteObjectName, server))
                    {
                        var thisAssembly = Assembly.GetEntryAssembly().Location;

                        Log("Launching client...");
                        using (var client = Process.Start(thisAssembly, "client"))
                        {
                            client.WaitForExit();
                        }
                    }
                }
                else if (args[0].Equals("client"))
                {
                    Log($"Client running...");

                    // Acquire remote object
                    object remoteObject = rot.GetObject(remoteObjectName);
                    var server = (IServer)remoteObject;

                    Log("Calling remote object...");
                    var pi = server.ComputePi();
                    Log($"\u03C0 = {pi}");

                    Log("Press any key to continue...");
                    Console.ReadKey();
                }
                else
                {
                    throw new ArgumentException("Invalid demo mode");
                }
            }
        }

        static void Log(string fmt, params object[] fmtargs)
        {
            Console.WriteLine($"Process {processId}: {string.Format(fmt, fmtargs)}");
        }
    }
}
