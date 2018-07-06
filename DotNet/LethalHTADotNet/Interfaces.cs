using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace LethalHTADotNet
{   
    public struct FILETIME
    {   
        public int dwLowDateTime;
        public int dwHighDateTime;
    }

    public struct ULARGE_INTEGER
    {
        public ulong QuadPart;
    }

    [ComImport]
    [Guid("0000010C-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersist
    {
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int GetClassID(out Guid pClassID);
    }

    [ComImport]
    [Guid("00000109-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistStream : IPersist
    {
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        new int GetClassID(out Guid pClassID);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int IsDirty();

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Load([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Save([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm, [In]  int fClearDirty);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void GetSizeMax([Out]  [MarshalAs(UnmanagedType.LPArray)] ULARGE_INTEGER[] pcbSize);
    }

    [ComImport]
    [Guid("0000000F-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]  
    public interface IMoniker : IPersistStream
    {
        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        new int GetClassID(out Guid pClassID);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        new int IsDirty();

        [MethodImpl(MethodImplOptions.InternalCall)]
        new void Load([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm);

        [MethodImpl(MethodImplOptions.InternalCall)]
        new void Save([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm, [In]  int fClearDirty);

        [MethodImpl(MethodImplOptions.InternalCall)]
        new void GetSizeMax([Out]  [MarshalAs(UnmanagedType.LPArray)] ULARGE_INTEGER[] pcbSize);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void BindToObject([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In]  ref Guid riidResult, [MarshalAs(UnmanagedType.IUnknown)] out object ppvResult);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void BindToStorage([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In]  ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Reduce([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In]  uint dwReduceHowFar, [In] [Out] [MarshalAs(UnmanagedType.Interface)] ref IMoniker ppmkToLeft, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkReduced);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void ComposeWith([In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkRight, [In]  int fOnlyIfNotGeneric, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkComposite);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Enum([In]  int fForward, [MarshalAs(UnmanagedType.Interface)] out IEnumMoniker ppenumMoniker);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void IsEqual([In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkOtherMoniker);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Hash( out uint pdwHash);

        [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall)]
        int IsRunning([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkNewlyRunning);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void GetTimeOfLastChange([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [Out]  [MarshalAs(UnmanagedType.LPArray)] FILETIME[] pFileTime);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void Inverse([MarshalAs(UnmanagedType.Interface)] out IMoniker ppmk);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void CommonPrefixWith([In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkOther, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkPrefix);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void RelativePathTo([In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkOther, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkRelPath);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void GetDisplayName([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft,  [MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplayName);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void ParseDisplayName([In] [MarshalAs(UnmanagedType.Interface)] IBindCtx pbc, [In] [MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In]  [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName,  out uint pchEaten, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkOut);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void IsSystemMoniker( out uint pdwMksys);
    }

    [ComImport]
    [Guid("00000003-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComConversionLoss]
    public interface IMarshal
    {
        [MethodImpl(MethodImplOptions.InternalCall)]
        void GetUnmarshalClass([In]  ref Guid riid, [In] IntPtr pv, [In]  uint dwDestContext, [In] IntPtr pvDestContext, [In]  uint MSHLFLAGS, out Guid pCid);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void GetMarshalSizeMax([In]  ref Guid riid, [In] IntPtr pv, [In]  uint dwDestContext, [In] IntPtr pvDestContext, [In]  uint MSHLFLAGS,  out uint pSize);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void MarshalInterface([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm, [In]  ref Guid riid, [In] IntPtr pv, [In]  uint dwDestContext, [In] IntPtr pvDestContext, [In]  uint MSHLFLAGS);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void UnmarshalInterface([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm, [In]  ref Guid riid, out IntPtr ppv);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void ReleaseMarshalData([In] [MarshalAs(UnmanagedType.Interface)] IStream pstm);

        [MethodImpl(MethodImplOptions.InternalCall)]
        void DisconnectObject([In]  uint dwReserved);
    }

    [Guid("79EAC9C9-BAF9-11CE-8C82-00AA004BA90B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPersistMoniker
    {
        void GetClassID(out Guid p0);
        void IsDirty();
        void Load(uint fFullyAvailable, LethalHTADotNet.IMoniker pimkName, IBindCtx pibc, uint grfMode);
        void Save(LethalHTADotNet.IMoniker pimkName, IBindCtx pbc, uint fRemember);
        void SaveCompleted(LethalHTADotNet.IMoniker pimkName, IBindCtx pibc);
        void GetCurMoniker(out LethalHTADotNet.IMoniker ppimkName);
    }
}
