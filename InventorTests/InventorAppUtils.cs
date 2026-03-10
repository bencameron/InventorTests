using System;
using System.Runtime.InteropServices;
using System.Security;
using Inventor;

public static class InventorAppUtils
{
    private static class NativeMethods
    {
        private const string OLEAUT32 = "oleaut32.dll";
        private const string OLE32 = "ole32.dll";

        [DllImport(OLE32, PreserveSig = false)]
        [SuppressUnmanagedCodeSecurity]
        internal static extern void CLSIDFromProgIDEx(
            [MarshalAs(UnmanagedType.LPWStr)] string progId,
            out Guid clsid);

        [DllImport(OLE32, PreserveSig = false)]
        [SuppressUnmanagedCodeSecurity]
        internal static extern void CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string progId,
            out Guid clsid);

        [DllImport(OLEAUT32, PreserveSig = false)]
        [SuppressUnmanagedCodeSecurity]
        internal static extern void GetActiveObject(
            ref Guid rclsid,
            IntPtr reserved,
            [MarshalAs(UnmanagedType.Interface)] out object ppunk);
    }

    public static Application GetInventorInstance(bool forceNewInstance = false)
    {
        Application app = null;

        if (!forceNewInstance)
        {
            try
            {
                Guid clsid;
                string progId = "Inventor.Application";

                try
                {
                    NativeMethods.CLSIDFromProgIDEx(progId, out clsid);
                }
                catch
                {
                    NativeMethods.CLSIDFromProgID(progId, out clsid);
                }

                NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                app = (Application)obj;
            }
            catch
            {
                app = GetNewInventorInstance();
            }
        }

        if (app == null)
        {
            app = GetNewInventorInstance();
        }

        return app;
    }

    private static Application GetNewInventorInstance()
    {
        var inventorType = Type.GetTypeFromProgID("Inventor.Application", true);
        return (Application)Activator.CreateInstance(inventorType);
    }
}