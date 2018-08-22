namespace LoaderShim
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;

    [ComImport]
    [ComVisible(false)]
    [Guid("00000001-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IClassFactory
    {
        void CreateInstance(
            [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

        void LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ActivationRequest
    {
        public Guid ClassId;
        public Guid InterfaceId;
        public int AssemblyCount;
        public IntPtr AssemblyList;
        public IntPtr ClassFactoryDest;
    }

    public class ComActivator
    {
        /// <summary>
        /// Entry point for unmanaged hosting API
        /// </summary>
        /// <param name="arg">String argument - address of <see cref="ActivationRequest"/> instance</param>
        /// <returns>Error code</returns>
        public static int GetClassFactoryForType(string arg)
        {
            try
            {
                var reqPtr = new IntPtr(Int64.Parse(arg));
                var req = Marshal.PtrToStructure<ActivationRequest>(reqPtr);

                if (req.InterfaceId != typeof(IClassFactory).GUID)
                {
                    throw new NotSupportedException();
                }

                var cf = new BasicClassFactory(req.ClassId);
                IntPtr nativeIUnknown = Marshal.GetIUnknownForObject(cf);
                Marshal.WriteIntPtr(req.ClassFactoryDest, nativeIUnknown);
            }
            catch (Exception e)
            {
                return e.HResult;
            }

            return 0;
        }

        [ComVisible(true)]
        internal class BasicClassFactory : IClassFactory
        {
            private readonly Guid ClassId;
            private readonly Type ClassType;
            private readonly Assembly ClassAssembly;

            public BasicClassFactory(Guid clsid)
            {
                this.ClassId = clsid;

                // [TODO] Determine what assembly the class is in
                this.ClassAssembly = Assembly.Load("NetServer");

                foreach (Type t in this.ClassAssembly.GetTypes())
                {
                    if (t.GUID == this.ClassId)
                    {
                        this.ClassType = t;
                        break;
                    }
                }

                if (this.ClassType == null)
                {
                    const int CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
                    throw new COMException(string.Empty, CLASS_E_CLASSNOTAVAILABLE);
                }
            }

            void IClassFactory.CreateInstance(
                [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
                ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppvObject)
            {
                // Verify the class implements the desired interface
                foreach (Type i in this.ClassType.GetInterfaces())
                {
                    if (i.GUID == riid)
                    {
                        ppvObject = Activator.CreateInstance(this.ClassType);

                        if (pUnkOuter != null)
                        {
                            try
                            {
                                IntPtr outerPtr = Marshal.GetIUnknownForObject(pUnkOuter);
                                IntPtr innerPtr = Marshal.CreateAggregatedObject(outerPtr, ppvObject);
                                ppvObject = Marshal.GetObjectForIUnknown(innerPtr);
                            }
                            finally
                            {
                                // Decrement the above 'Marshal.GetIUnknownForObject()'
                                Marshal.ReleaseComObject(pUnkOuter);
                            }
                        }

                        return;
                    }
                }

                // E_NOINTERFACE
                throw new InvalidCastException();
            }

            void IClassFactory.LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock)
            {
                // nop
            }
        }
    }
}
