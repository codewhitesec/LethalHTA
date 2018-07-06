using System;
using System.Runtime.InteropServices;

namespace LethalHTADotNet
{
    [ComVisible(true)]
    class FakeObject : IMarshal, IMoniker
    {
        private IMarshal _marshal;

        public FakeObject(IMoniker moniker)
        {
            this._marshal = (IMarshal)moniker;
        }

        public void GetUnmarshalClass([In] ref Guid riid, [In] IntPtr pv, [In] uint dwDestContext, [In] IntPtr pvDestContext, [In] uint MSHLFLAGS, out Guid pCid)
        {
            _marshal.GetUnmarshalClass(riid, pv, 1, pvDestContext, MSHLFLAGS, out pCid);
        }

        public void GetMarshalSizeMax([In] ref Guid riid, [In] IntPtr pv, [In] uint dwDestContext, [In] IntPtr pvDestContext, [In] uint MSHLFLAGS, out uint pSize)
        {
            _marshal.GetMarshalSizeMax(riid, pv, 1, pvDestContext, MSHLFLAGS, out pSize);
        }

        public void MarshalInterface([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm, [In] ref Guid riid, [In] IntPtr pv, [In] uint dwDestContext, [In] IntPtr pvDestContext, [In] uint MSHLFLAGS)
        {
            _marshal.MarshalInterface(pstm, riid, pv, 1, pvDestContext, MSHLFLAGS);
        }

        public void UnmarshalInterface([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm, [In] ref Guid riid, out IntPtr ppv)
        {
            _marshal.UnmarshalInterface(pstm, ref riid, out ppv);
        }

        public void ReleaseMarshalData([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm)
        {
            _marshal.ReleaseMarshalData(pstm);
        }

        public void DisconnectObject([In] uint dwReserved)
        {
            _marshal.DisconnectObject(dwReserved);
        }


        public int GetClassID(out Guid pClassID)
        {
            throw new NotImplementedException();
        }

        public int IsDirty()
        {
            throw new NotImplementedException();
        }

        public void Load([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm)
        {
            throw new NotImplementedException();
        }

        public void Save([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm, [In] int fClearDirty)
        {
            throw new NotImplementedException();
        }

        public void GetSizeMax([MarshalAs(UnmanagedType.LPArray), Out] ULARGE_INTEGER[] pcbSize)
        {
            throw new NotImplementedException();
        }

        public void BindToObject([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In] ref Guid riidResult, [MarshalAs(UnmanagedType.IUnknown)] out object ppvResult)
        {
            throw new NotImplementedException();
        }

        public void BindToStorage([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In] ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj)
        {
            throw new NotImplementedException();
        }

        public void Reduce([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In] uint dwReduceHowFar, [In, MarshalAs(UnmanagedType.Interface), Out] ref IMoniker ppmkToLeft, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkReduced)
        {
            throw new NotImplementedException();
        }

        public void ComposeWith([In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkRight, [In] int fOnlyIfNotGeneric, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkComposite)
        {
            throw new NotImplementedException();
        }

        public void Enum([In] int fForward, [MarshalAs(UnmanagedType.Interface)] out System.Runtime.InteropServices.ComTypes.IEnumMoniker ppenumMoniker)
        {
            throw new NotImplementedException();
        }

        public void IsEqual([In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkOtherMoniker)
        {
            throw new NotImplementedException();
        }

        public void Hash(out uint pdwHash)
        {
            throw new NotImplementedException();
        }

        public int IsRunning([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkNewlyRunning)
        {
            throw new NotImplementedException();
        }

        public void GetTimeOfLastChange([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [MarshalAs(UnmanagedType.LPArray), Out] FILETIME[] pFileTime)
        {
            throw new NotImplementedException();
        }

        public void Inverse([MarshalAs(UnmanagedType.Interface)] out IMoniker ppmk)
        {
            throw new NotImplementedException();
        }

        public void CommonPrefixWith([In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkOther, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkPrefix)
        {
            throw new NotImplementedException();
        }

        public void RelativePathTo([In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkOther, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkRelPath)
        {
            throw new NotImplementedException();
        }

        public void GetDisplayName([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplayName)
        {
            throw new NotImplementedException();
        }

        public void ParseDisplayName([In, MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [In, MarshalAs(UnmanagedType.Interface)] IMoniker pmkToLeft, [In, MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, out uint pchEaten, [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmkOut)
        {
            throw new NotImplementedException();
        }

        public void IsSystemMoniker(out uint pdwMksys)
        {
            throw new NotImplementedException();
        }


    }
}
