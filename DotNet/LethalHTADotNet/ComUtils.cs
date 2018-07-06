using System;

namespace LethalHTADotNet
{
    public static class ComUtils
    {
        public static IntPtr IID_IUnknownPtr = GuidToPointer("00000000-0000-0000-C000-000000000046");

        public static IntPtr GuidToPointer(string guid)
        {
            Guid g = new Guid(guid);

            IntPtr ret = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(16);
            System.Runtime.InteropServices.Marshal.Copy(g.ToByteArray(), 0, ret, 16);

            return ret;
        }

        [Flags]
        public enum CLSCTX : uint
        {
            CLSCTX_INPROC_SERVER = 0x1,
            CLSCTX_INPROC_HANDLER = 0x2,
            CLSCTX_LOCAL_SERVER = 0x4,
            CLSCTX_INPROC_SERVER16 = 0x8,
            CLSCTX_REMOTE_SERVER = 0x10,
            CLSCTX_INPROC_HANDLER16 = 0x20,
            CLSCTX_RESERVED1 = 0x40,
            CLSCTX_RESERVED2 = 0x80,
            CLSCTX_RESERVED3 = 0x100,
            CLSCTX_RESERVED4 = 0x200,
            CLSCTX_NO_CODE_DOWNLOAD = 0x400,
            CLSCTX_RESERVED5 = 0x800,
            CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
            CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
            CLSCTX_NO_FAILURE_LOG = 0x4000,
            CLSCTX_DISABLE_AAA = 0x8000,
            CLSCTX_ENABLE_AAA = 0x10000,
            CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000,
            CLSCTX_ACTIVATE_32_BIT_SERVER = 0x40000,
            CLSCTX_ACTIVATE_64_BIT_SERVER = 0x80000,
            CLSCTX_INPROC = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
            CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,
            CLSCTX_ALL = CLSCTX_SERVER | CLSCTX_INPROC_HANDLER
        }

        [System.Runtime.InteropServices.DllImport("urlmon.dll")]
        public static extern int CreateURLMonikerEx(
                IntPtr punk,
                [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszDisplayName,
                out IMoniker ppmk,
                uint flags
            );

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct MULTI_QI
        {
            public IntPtr pIID;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Interface)]
            public object pItf;
            public int hr;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public class COSERVERINFO
        {
            public uint dwReserved1;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)]
            public string pwszName;
            public IntPtr pAuthInfo;
            public uint dwReserved2;
        }

        [System.Runtime.InteropServices.DllImport("ole32.dll")]
        public static extern void CoCreateInstanceEx(
           [System.Runtime.InteropServices.In, System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStruct)] Guid rclsid,
           [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.IUnknown)] object pUnkOuter,
           CLSCTX dwClsCtx,
           COSERVERINFO pServerInfo,
           uint cmq,
           [System.Runtime.InteropServices.In, System.Runtime.InteropServices.Out] MULTI_QI[] pResults);
    }
}
