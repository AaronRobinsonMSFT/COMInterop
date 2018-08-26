namespace CoreShim
{
    using System;
    using System.Diagnostics;
    using System.IO;
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
    internal struct LICINFO
    {
        public int cbLicInfo;

        [MarshalAs(UnmanagedType.Bool)]
        public bool fRuntimeKeyAvail;

        [MarshalAs(UnmanagedType.Bool)]
        public bool fLicVerified;
    }

    [ComImport]
    [ComVisible(false)]
    [Guid("B196B28F-BAB4-101A-B69C-00AA00341D07")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IClassFactory2 : IClassFactory
    {
        new void CreateInstance(
            [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

        new void LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);

        void GetLicInfo(ref LICINFO pLicInfo);

        void RequestLicKey(
            int dwReserved,
            [MarshalAs(UnmanagedType.BStr)] out string pBstrKey);

        void CreateInstanceLic(
            [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
            [MarshalAs(UnmanagedType.Interface)] object pUnkReserved,
            ref Guid riid,
            [MarshalAs(UnmanagedType.BStr)] string bstrKey,
            [MarshalAs(UnmanagedType.Interface)] out object ppvObject);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ComActivationContext
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
        /// Entry point for unmanaged COM activation API
        /// </summary>
        /// <param name="cxt">Reference to a <see cref="ComActivationContext"/> instance</param>
        /// <returns>Error code</returns>
        public static int GetClassFactoryForType(ref ComActivationContext cxt)
        {
            Debug.WriteLine(
$@"{nameof(GetClassFactoryForType)} arguments:
    {cxt.ClassId}
    {cxt.InterfaceId}
    {cxt.AssemblyCount}
    0x{cxt.AssemblyList.ToInt64():x}
    0x{cxt.ClassFactoryDest.ToInt64():x}");

            try
            {
                if (cxt.InterfaceId != typeof(IClassFactory).GUID
                    && cxt.InterfaceId != typeof(IClassFactory2).GUID)
                {
                    throw new NotSupportedException();
                }

                string[] potentialAssembies = CreateAssemblyArray(cxt.AssemblyCount, cxt.AssemblyList);
                (Assembly classAssembly, Type classType) = FindClassAssemblyAndType(cxt.ClassId, potentialAssembies);

                var cf = new BasicClassFactory(cxt.ClassId, classAssembly, classType);
                IntPtr nativeIUnknown = Marshal.GetIUnknownForObject(cf);
                Marshal.WriteIntPtr(cxt.ClassFactoryDest, nativeIUnknown);
            }
            catch (Exception e)
            {
                return e.HResult;
            }

            return 0;
        }

        private static string[] CreateAssemblyArray(int assemblyCount, IntPtr assemblyList)
        {
            var assemblies = new string[assemblyCount];

            unsafe
            {
                var spanOfPtrs = new Span<IntPtr>(assemblyList.ToPointer(), assemblyCount);
                for (int i = 0; i < assemblyCount; ++i)
                {
                    assemblies[i] = Marshal.PtrToStringUni(spanOfPtrs[i]);
                }
            }

            return assemblies;
        }

        private static (Assembly assembly, Type type) FindClassAssemblyAndType(Guid clsid, string[] potentialAssembies)
        {
            // Determine what assembly the class is in
            foreach (string assemPath in potentialAssembies)
            {
                Assembly assem;
                try
                {
                    string assemPathLocal = assemPath;
                    string extMaybe = Path.GetExtension(assemPath);
                    if (".manifest".Equals(extMaybe, StringComparison.OrdinalIgnoreCase))
                    {
                        assemPathLocal = Path.ChangeExtension(assemPath, ".dll");
                    }

                    assem = Assembly.LoadFrom(assemPathLocal);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e);
                    continue;
                }

                // Check the loaded assembly for a class with the desired ID
                foreach (Type t in assem.GetTypes())
                {
                    if (t.GUID == clsid)
                    {
                        return (assem, t);
                    }
                }
            }

            const int CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
            throw new COMException(string.Empty, CLASS_E_CLASSNOTAVAILABLE);
        }

        [ComVisible(true)]
        internal class BasicClassFactory : IClassFactory2
        {
            private static readonly Guid Clsid_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");
            private readonly Guid classId;
            private readonly Type classType;
            private readonly Assembly classAssembly;

            public BasicClassFactory(Guid clsid, Assembly assembly, Type classType)
            {
                this.classId = clsid;
                this.classType = classType;
                this.classAssembly = assembly;
            }

            public void CreateInstance(
                [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
                ref Guid riid,
                [MarshalAs(UnmanagedType.Interface)] out object ppvObject)
            {
                if (riid != Clsid_IUnknown)
                {
                    bool found = false;

                    // Verify the class implements the desired interface
                    foreach (Type i in this.classType.GetInterfaces())
                    {
                        if (i.GUID == riid)
                        {
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        // E_NOINTERFACE
                        throw new InvalidCastException();
                    }
                }

                ppvObject = Activator.CreateInstance(this.classType);
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
            }

            public void LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock)
            {
                // nop
            }

            public void GetLicInfo(ref LICINFO pLicInfo)
            {
                throw new NotImplementedException();
            }

            public void RequestLicKey(int dwReserved, [MarshalAs(UnmanagedType.BStr)] out string pBstrKey)
            {
                throw new NotImplementedException();
            }

            public void CreateInstanceLic(
                [MarshalAs(UnmanagedType.Interface)] object pUnkOuter,
                [MarshalAs(UnmanagedType.Interface)] object pUnkReserved,
                ref Guid riid,
                [MarshalAs(UnmanagedType.BStr)] string bstrKey,
                [MarshalAs(UnmanagedType.Interface)] out object ppvObject)
            {
                throw new NotImplementedException();
            }
        }
    }
}
