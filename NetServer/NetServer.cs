namespace NetServer
{
    using System;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("F38720E5-2D64-445E-88FB-1D696F614C78")]
    public interface IServer
    {
        double ComputePi();
    }

    [ComVisible(true)]
    [Guid("114383E9-1969-47D2-9AA9-91388C961A19")]
    public class Server : IServer
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
