﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Vanara.Extensions;
using Vanara.PInvoke;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDataObject = System.Windows.Forms.IDataObject;



namespace ClipSharp
{
    [ComVisible(true)]
    public class DataStore : IComDataObject, IDataObject
    {
        private readonly Dictionary<FormatId, Dictionary<int,object>> store = new Dictionary<FormatId, Dictionary<int, object>>();

        private const TYMED AcceptableTymed = TYMED.TYMED_ENHMF | TYMED.TYMED_HGLOBAL | TYMED.TYMED_ISTORAGE |
                                              TYMED.TYMED_MFPICT | TYMED.TYMED_GDI;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool IsAcceptableTymed(TYMED val) => (val & AcceptableTymed) != 0;

        public DataStore()
        {
        }

        public void SetData<T>(T data, int lindex = -1)
        {
            SetData(FormatId.FromName(typeof(T).FullName), data);
        }

        public void SetData(FormatId id, object data , int lindex = -1)
        {
            if(!store.TryGetValue(id,out var dic) || dic == null)
            {
                store[id] = new Dictionary<int, object>();
            }
            store[id][lindex] = data;
        }

        public void SetString(string str, NativeStringType native = NativeStringType.Unicode)
        {
            switch (native)
            {
                case NativeStringType.Unicode:
                    SetData(FormatId.CF_UNICODETEXT,str);
                    break;
                case NativeStringType.Ansi:
                    SetData(FormatId.CF_TEXT,str);
                    break;
                default:
                    break;
            }
        }

        public void SetFileContents(Dictionary<FileDescriptor,Stream> contents)
        {
            SetData(FormatId.CFSTR_FILEDESCRIPTORW, FileDescriptor.CreateNativeFileDescriptors(contents.Keys ));

            int i = 0;
            foreach (var (des,st) in contents)
            {
                SetData(FormatId.CFSTR_FILECONTENTS, st, i++);
            }
        }


        public object GetData(FormatId id,int lindex = -1)
        {
            if (store.TryGetValue(id, out var dict) && dict.TryGetValue(lindex, out var obj))
                return obj;
            throw new KeyNotFoundException();
        }

        public T GetData<T>(int lindex = -1)
        {
            return GetData<T>(FormatId.FromName(typeof(T).FullName), lindex);
        }

        public T GetData<T>(FormatId id, int lindex = -1)
        {
            if (store.TryGetValue(id, out var dict) && dict.TryGetValue(lindex, out var obj))
                return (T)obj;
            throw new KeyNotFoundException();
        }

        public bool GetDataPresent(FormatId id,int lindex = -1)
        {
            return store.ContainsKey(id) && (store[id]?.ContainsKey(lindex) ?? false);
        }
        public bool TryGetData<T>(FormatId id, out T data, int lindex = -1)
        {
            if (store.TryGetValue(id, out var dict))
                if (dict.TryGetValue(lindex, out var obj))
                    if (obj is T d)
                    {
                        data = d;
                        return true;
                    }

            data = default;
            return false;
        }

        public IEnumerable<FormatId> GetFormats() => store.Keys;


        #region System.Windows.Forms.IComDataObject
        object IDataObject.GetData(string format, bool autoConvert) => GetData<object>(FormatId.FromName("format"));

        object IDataObject.GetData(string format) => GetData<object>(FormatId.FromName(format));

        object IDataObject.GetData(Type format) => GetData<object>(FormatId.FromName(format.FullName));

        void IDataObject.SetData(string format, bool autoConvert, object data) => SetData(FormatId.FromName(format), data);

        void IDataObject.SetData(string format, object data) => ((IDataObject)this).SetData(format, false, data);

        void IDataObject.SetData(Type format, object data) => ((IDataObject)this).SetData(format.FullName, false, data);

        void IDataObject.SetData(object data) => ((IDataObject)this).SetData(data.GetType().FullName, false, data);

        bool IDataObject.GetDataPresent(string format, bool autoConvert) =>
            GetDataPresent(FormatId.FromName(format));

        bool IDataObject.GetDataPresent(string format) => ((IDataObject)this).GetDataPresent(format, false);

        bool IDataObject.GetDataPresent(Type format) => ((IDataObject)this).GetDataPresent(format.FullName, false);


        string[] IDataObject.GetFormats(bool autoConvert) => this.GetFormats().Select(x => x.DotNetName).ToArray();

        string[] IDataObject.GetFormats() => this.GetFormats().Select(x => x.DotNetName).ToArray();
        #endregion

        #region IComDataObject
        int IComDataObject.DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            connection = 0;
            return HRESULT.E_NOTIMPL.Code;
        }

        void IComDataObject.DUnadvise(int connection) => HRESULT.E_NOTIMPL.ThrowIfFailed();

        int IComDataObject.EnumDAdvise(out IEnumSTATDATA? enumAdvise)
        {
            enumAdvise = null;
            return unchecked((int)0x80040003); // unchecked((int)0x80040003),
        }

        IEnumFORMATETC IComDataObject.EnumFormatEtc(DATADIR direction)
        {
            if (direction == DATADIR.DATADIR_GET)
            {
                return new FormatEnumrator(this);
            }
            throw HRESULT.E_NOTIMPL.GetException();
        }

        int IComDataObject.GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            formatOut = default;
            return 0x00040130; //DATA_S_SAMEFORMATETC
        }
        void IComDataObject.GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            medium = new STGMEDIUM();
            var id = format.GetFormatId();
            if (!TryGetData(id, out object val,format.lindex))
            {
                Marshal.ThrowExceptionForHR(unchecked((int)0x80040064));
            }

            var ptr = medium.unionmember;
            if (id == FormatId.CF_HDROP && val is string[] files)
            {
                SaveHdropToHandle(ref ptr, files);
                medium.unionmember = ptr;
                return;
            }

            if (id == FormatId.CF_BITMAP && val is Bitmap bmp)
            {
                medium.unionmember = GetCompatibleBitmap(bmp).DangerousGetHandle();
                return;
            }



            if ((format.tymed & TYMED.TYMED_HGLOBAL) != 0)
            {
                medium.tymed = TYMED.TYMED_HGLOBAL;
                var hglobal = Kernel32.GlobalAlloc(Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT, 1);
                if (hglobal.IsNull)
                {
                    throw new OutOfMemoryException();
                }
                medium.unionmember = hglobal.DangerousGetHandle();
                try
                {
                    ((IComDataObject)this).GetDataHere(ref format, ref medium);
                }
                catch
                {
                    Kernel32.GlobalFree(hglobal);
                }
            }
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern unsafe int WideCharToMultiByte(
            uint CodePage,
            Kernel32.WCCONV dwFlags,
            [MarshalAs(UnmanagedType.LPWStr), In] string lpWideCharStr,
            int cchWideChar,
            byte* lpMultiByteStr,
            int cbMultiByte,
            IntPtr lpDefaultChar = default(IntPtr),
            IntPtr lpUsedDefaultChar = default(IntPtr));
        internal static unsafe int StringToAnsiString(string s, byte* buffer, int bufferLength, bool bestFit = false, bool throwOnUnmappableChar = false)
        {
            int nb;

            var flags = bestFit ? (Kernel32.WCCONV)0 : Kernel32.WCCONV.WC_NO_BEST_FIT_CHARS;
            uint defaultCharUsed = 0;

            nb = WideCharToMultiByte(
                Kernel32.CP_ACP,
                flags,
                s,
                s.Length,
                buffer,
                bufferLength,
                IntPtr.Zero,
                throwOnUnmappableChar ? new IntPtr(&defaultCharUsed) : IntPtr.Zero);

            if (defaultCharUsed != 0)
            {
                throw new ArgumentException(nameof(defaultCharUsed));
            }

            buffer[nb] = 0;
            return nb;
        }

        private unsafe HRESULT SaveStringToHandle(ref IntPtr handle, string s, NativeStringType type)
        {
            if (handle == IntPtr.Zero) return HRESULT.E_INVALIDARG;
            switch (type)
            {
                case NativeStringType.Unicode:
                    {
                        int nb = (s.Length + 1) * 2;

                        // Overflow checking
                        if (nb < s.Length) throw new ArgumentOutOfRangeException(nameof(s));

                        var hg = Kernel32.GlobalReAlloc(handle, nb,
                            Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT);
                        handle = hg.DangerousGetHandle();
                        if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;
                        var lc = Kernel32.GlobalLock(hg);

                        fixed (char* firstChar = s)
                        {
                            Buffer.MemoryCopy(firstChar, lc.ToPointer(), nb, nb);
                        }

                        Kernel32.GlobalUnlock(hg);
                        break;
                    }
                case NativeStringType.Ansi:
                    {
                        // Ansi. See also StringToHGlobalAnsi

                        long lnb = (s.Length + 1) * (long)Marshal.SystemMaxDBCSCharSize;
                        int nb = (int)lnb;
                        if (nb != lnb) throw new ArgumentOutOfRangeException(nameof(s));

                        var hg = Kernel32.GlobalReAlloc(handle, nb,
                            Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT);
                        handle = hg.DangerousGetHandle();
                        if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;

                        var lc = Kernel32.GlobalLock(hg);
                        StringToAnsiString(s, (byte*)lc.ToPointer(), nb);
                        Kernel32.GlobalUnlock(hg);
                        break;
                    }
                case NativeStringType.Utf8:
                    {
                        var enc = new UTF8Encoding();
                        int nb = enc.GetByteCount(s) + 1;
                        var hg = Kernel32.GlobalReAlloc(handle, nb,
                            Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT);
                        handle = hg.DangerousGetHandle();
                        if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;

                        var lc = Kernel32.GlobalLock(hg);
                        fixed (char* firstChar = s)
                        {
                            enc.GetBytes(firstChar, s.Length, (byte*)lc.ToPointer(), nb);
                        }
                        Kernel32.GlobalUnlock(hg);
                        break;
                    }
            }
            return HRESULT.S_OK;
        }

        unsafe HRESULT SaveHdropToHandle(ref IntPtr handle, string[] files)
        {
            int size = sizeof(Shell32.DROPFILES);
            foreach (var f in files)
            {
                size += (f.Length + 1) * 2;
            }

            size += 2; // Terminal null

            var hg = Kernel32.GlobalReAlloc(handle, size, Kernel32.GMEM.GMEM_MOVEABLE);
            if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;

            var lc = Kernel32.GlobalLock(hg);
            Shell32.DROPFILES* df = (Shell32.DROPFILES*)lc.ToPointer();
            df->pFiles = (uint)files.Length;
            df->fNC = false;
            df->fWide = false;

            char* ptr = ((char*)df) + sizeof(Shell32.DROPFILES);
            foreach (var f in files)
            {
                fixed (char* str = f)
                {
                    Buffer.MemoryCopy(str, ptr, f.Length + 2, f.Length + 2);
                    ptr[f.Length] = '\0';
                    ptr += f.Length + 1;
                }
                *ptr = '\0';
            }
            Kernel32.GlobalUnlock(hg);
            return HRESULT.S_OK;
        }

        unsafe HRESULT SaveSpanToHandle(ref IntPtr handle, Span<byte> src)
        {
            var hg = Kernel32.GlobalReAlloc(handle, src.Length, Kernel32.GMEM.GMEM_MOVEABLE);
            if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;
            handle = hg.DangerousGetHandle();
            var dist = new Span<byte>(Kernel32.GlobalLock(hg).ToPointer(), src.Length);
            src.CopyTo(dist);
            Kernel32.GlobalUnlock(hg);
            return HRESULT.S_OK;

        }

        unsafe HRESULT SaveStreamToHandle(ref IntPtr handle, Stream src)
        {
            var hg = Kernel32.GlobalReAlloc(handle, src.Length, Kernel32.GMEM.GMEM_MOVEABLE);
            if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;
            handle = hg.DangerousGetHandle();
            var dist = new Span<byte>(Kernel32.GlobalLock(hg).ToPointer(), (int)src.Length);
            src.Position = 0;
            src.Read(dist);
            Kernel32.GlobalUnlock(hg);
            return HRESULT.S_OK;

        }

        [DllImport("gdi32.dll")]
        public static extern HBITMAP CreateCompatibleBitmap(HDC hdc, int cx, int cy);

        private HBITMAP GetCompatibleBitmap(Bitmap bm)
        {
            using var hDC = User32.GetDC(IntPtr.Zero);

            // GDI+ returns a DIBSECTION based HBITMAP. The clipboard deals well
            // only with bitmaps created using CreateCompatibleBitmap(). So, we
            // convert the DIBSECTION into a compatible bitmap.
            IntPtr hBitmap = bm.GetHbitmap();

            // Create a compatible DC to render the source bitmap.
            using var dcSrc = Gdi32.CreateCompatibleDC(hDC);
            var srcOld = Gdi32.SelectObject(dcSrc, hBitmap);

            // Create a compatible DC and a new compatible bitmap.
            using var dcDest = Gdi32.CreateCompatibleDC(hDC);
            var hBitmapNew = CreateCompatibleBitmap(hDC, bm.Size.Width, bm.Size.Height);

            // Select the new bitmap into a compatible DC and render the blt the original bitmap.
            var destOld = Gdi32.SelectObject(dcDest, hBitmapNew);
            Gdi32.BitBlt(
                dcDest,
                0,
                0,
                bm.Size.Width,
                bm.Size.Height,
                dcSrc,
                0,
                0,
                Gdi32.RasterOperationMode.SRCCOPY);

            // Clear the source and destination compatible DCs.
            Gdi32.SelectObject(dcSrc, srcOld);
            Gdi32.SelectObject(dcDest, destOld);

            return hBitmapNew;
        }


        void IComDataObject.GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            var id = new FormatId(format.cfFormat);
            if (!TryGetData(id, out object val,format.lindex))
            {
                //Marshal.ThrowExceptionForHR(unchecked((int)0x80040064));
                HRESULT.E_FAIL.ThrowIfFailed();
            }

            var ptr = medium.unionmember;
            if (id == FormatId.CF_HDROP && val is string[] files)
            {
                SaveHdropToHandle(ref ptr, files);
                medium.unionmember = ptr;
                return;
            }
            
            if (id == FormatId.CF_BITMAP && val is Bitmap bmp)
            {
                medium.unionmember = GetCompatibleBitmap(bmp).DangerousGetHandle();
                return;
            }

            (val switch
            {
                string s when (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                               id == FormatId.CommaSeparatedValue || id == FormatId.CFSTR_INETURLA)
                => SaveStringToHandle(ref ptr, s, NativeStringType.Ansi),
                string s when id == FormatId.Html || id == FormatId.Xaml
                => SaveStringToHandle(ref ptr, s, NativeStringType.Utf8),
                string s when (id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust ||
                               id == FormatId.CFSTR_INETURLW)
                => SaveStringToHandle(ref ptr, s, NativeStringType.Unicode),
                Memory<byte> m => SaveSpanToHandle(ref ptr, m.Span),
                byte[] b => SaveSpanToHandle(ref ptr, b.AsSpan()),
                Stream s => SaveStreamToHandle(ref ptr, s),
                _ => throw new ApplicationException()
            }).ThrowIfFailed();
            medium.unionmember = ptr;
        }

        private const int DV_E_TYMED = unchecked((int)0x80040069);
        private const int DV_E_DVASPECT = unchecked((int)0x8004006B);
        private const int DV_E_FORMATETC = unchecked((int)0x80040064);


        int IComDataObject.QueryGetData(ref FORMATETC format)
        {
            if (format.dwAspect != DVASPECT.DVASPECT_CONTENT) return DV_E_DVASPECT;
            if (!IsAcceptableTymed((format.tymed))) return DV_E_TYMED;
            if (format.cfFormat == 0) return HRESULT.S_FALSE.Code;
            if (!GetDataPresent(format.GetFormatId())) return DV_E_FORMATETC;
            return HRESULT.S_OK.Code;
        }

        void IComDataObject.SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

    class DataStoreEntry<T>
    {
        public T Data { get; }

        public DataStoreEntry(T data)
        {
            this.Data = data;
        }
    }

    [ComVisible(true)]
    public class FormatEnumrator : IEnumFORMATETC
    {
        private IEnumerator<FormatId> data;
        private int current = 0;
        public FormatEnumrator(DataStore dataObject)
        {
            var l = dataObject.GetFormats().ToList();
            this.data = l.GetEnumerator();
        }
        public void Clone(out IEnumFORMATETC newEnum)
        {

            throw new NotImplementedException();
        }

        public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
        {
            if (!data.MoveNext())
            {
                if (pceltFetched != null) pceltFetched[0] = 0;
                return HRESULT.S_FALSE.Code;
            }


            var id = data.Current.FormatEtc.GetFormatId();
            rgelt[0].cfFormat = (short)id.Id;

            if (id == FormatId.CF_BITMAP) rgelt[0].tymed = TYMED.TYMED_GDI;
            else if (id == FormatId.CF_ENHMETAFILE) rgelt[0].tymed = TYMED.TYMED_ENHMF;
            else if (id == FormatId.CF_METAFILEPICT) rgelt[0].tymed = TYMED.TYMED_MFPICT;
            else rgelt[0].tymed = TYMED.TYMED_HGLOBAL;
            rgelt[0].dwAspect = DVASPECT.DVASPECT_CONTENT;
            rgelt[0].ptd = IntPtr.Zero;
            rgelt[0].lindex = -1;
            if (pceltFetched != null) pceltFetched[0] = 1;

            return HRESULT.S_OK.Code;
        }

        public int Reset()
        {
            data.Reset();
            return HRESULT.S_OK.Code;
        }

        public int Skip(int celt)
        {
            do
            {
                if (!data.MoveNext()) return HRESULT.S_FALSE.Code;
            } while (celt-- > 0);

            return HRESULT.S_OK.Code;
        }
    }
}
