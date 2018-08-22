namespace NetServer
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Managed definition of COM interface
    /// </summary>
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("575375A2-3F92-44A8-89A3-A7BB87BE9622")]
    public interface IOuter
    {
        int ComputeFibonacci(int n);
    }

#pragma warning disable IDE1006 // Naming Styles
    /// <summary>
    /// Managed definition of CoClass
    /// </summary>
    [ComImport]
    [CoClass(typeof(OuterClass))]
    [Guid("575375A2-3F92-44A8-89A3-A7BB87BE9622")]
    public interface Outer : IOuter
    {
    }
#pragma warning restore IDE1006 // Naming Styles

    /// <summary>
    /// Managed activation for CoClass
    /// </summary>
    [ComImport]
    [Guid("BCF98F86-300D-4245-82F2-3D9D1E62B1FC")]
    public class OuterClass
    {
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F38720E5-2D64-445E-88FB-1D696F614C78")]
    public interface IServer
    {
        double ComputePi();
    }

    [ComVisible(true)]
    [Guid("114383E9-1969-47D2-9AA9-91388C961A19")]
    public class Server : OuterClass, IServer
    {
        public double ComputePi()
        {
            double sum = 0.0;
            int sign = 1;
            for (int i = 0; i < 1024; ++i)
            {
                sum += sign / (2.0 * i + 1.0);
                sign *= -1;
            }

            return 4.0 * sum;
        }
    }
}
