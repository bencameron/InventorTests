using Inventor;
using System.Runtime.InteropServices;

namespace InventorTests
{
    public static class InventorAppUtils
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
                                            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        static extern int CLSIDFromProgID(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        /// <summary>
        /// Gets the currently-running instance of Inventor or starts a new one if one is not already running
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static Application GetInventorInstance(bool forceNewInstance = false)
        {
            Application app = null;

            if (!forceNewInstance)
            {
                //Check for an instance of inventor, if one is not found then create one.
                try
                {
                    CLSIDFromProgID("Inventor.Application", out var classId);
                    GetActiveObject(ref classId, default, out var appObject);
                    app = (Application)appObject;
                }
                catch (Exception)
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
            var inventorAppType = Type.GetTypeFromProgID("Inventor.Application", true);
            return (Application)Activator.CreateInstance(inventorAppType);
        }
    }
}
