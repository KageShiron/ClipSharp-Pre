using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Vanara.PInvoke;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDataObject = System.Windows.Forms.IDataObject;



namespace ClipSharp
{
    public class DataObject : IComDataObject, IDataObject
    {
        private readonly IDataObject innerData;
        private readonly Dictionary<FormatId, object> store = new Dictionary<FormatId, object>();

        public DataObject(IDataObject data)
        {
            innerData = data;
        }

        public DataObject(IComDataObject data)
        {
            if (data is IDataObject dataObject)
            {
                innerData = dataObject;
            }
            else
            {
                innerData = new OleDataObject(data);
            }
        }

        public void SetData<T>(T data)
        {
            SetData(FormatId.FromDotNetName(typeof(T).FullName), data);
        }

        public void SetData(FormatId id, object data)
        {
            store[id] = data;
        }

        public object GetData(FormatId id)
        {
            return store.TryGetValue(id, out object value) ? value : throw new DirectoryNotFoundException();
        }

        public T GetData<T>()
        {
            return GetData<T>(FormatId.FromDotNetName(typeof(T).FullName));
        }

        public T GetData<T>(FormatId id)
        {
            return store.TryGetValue(id, out object value) ? (T)value : throw new DirectoryNotFoundException();
        }

        public bool GetDataPresent(FormatId id)
        {
            return store.ContainsKey(id);
        }
        public bool TryGetData<T>(FormatId id, out T data)
        {
            if (store.TryGetValue(id, out object value))
            {
                if (value is T d)
                {
                    data = d;
                    return true;
                }
            }

            data = default;
            return false;
        }

        public IEnumerable<FormatId> GetFormats() => store.Keys;


        #region System.Windows.Forms.IComDataObject
        object IDataObject.GetData(string format, bool autoConvert) => GetData<object>(FormatId.FromDotNetName("format"));

        object IDataObject.GetData(string format) => GetData<object>(FormatId.FromDotNetName(format));

        object IDataObject.GetData(Type format) => GetData<object>(FormatId.FromDotNetName(format.FullName));

        void IDataObject.SetData(string format, bool autoConvert, object data) => SetData(FormatId.FromDotNetName(format), data);

        void IDataObject.SetData(string format, object data) => ((IDataObject)this).SetData(format, false, data);

        void IDataObject.SetData(Type format, object data) => ((IDataObject)this).SetData(format.FullName, false, data);

        void IDataObject.SetData(object data) => ((IDataObject)this).SetData(data.GetType().FullName, false, data);

        bool IDataObject.GetDataPresent(string format, bool autoConvert) =>
            GetDataPresent(FormatId.FromDotNetName(format));

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
                return new FormatEnumerator(this);
            }
            else
            {
                HRESULT.E_NOTIMPL.ThrowIfFailed();
            }
        }

        int IComDataObject.GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            formatOut = default;
            return 0x00040130; //DATA_S_SAMEFORMATETC
        }
        void IComDataObject.GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            medium = new STGMEDIUM();
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

        private unsafe HRESULT SaveStringToHandle(ref IntPtr handle, string s, bool utf16)
        {
            if (handle == IntPtr.Zero) return HRESULT.E_INVALIDARG;
            if (utf16)
            {
                int nb = (s.Length + 1) * 2;

                // Overflow checking
                if (nb < s.Length) throw new ArgumentOutOfRangeException(nameof(s));

                var hg = Kernel32.GlobalReAlloc(handle, nb, Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT);
                handle = hg.DangerousGetHandle();
                if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;
                var lc = Kernel32.GlobalLock(hg);

                fixed (char* firstChar = s)
                {
                    Buffer.MemoryCopy(firstChar, lc.ToPointer(), nb, nb);
                }
                Kernel32.GlobalUnlock(hg);
            }
            else
            {
                // Ansi. See also StringToHGlobalAnsi

                long lnb = (s.Length + 1) * (long)Marshal.SystemMaxDBCSCharSize;
                int nb = (int)lnb;
                if (nb != lnb) throw new ArgumentOutOfRangeException(nameof(s));

                var hg = Kernel32.GlobalReAlloc(handle, nb, Kernel32.GMEM.GMEM_MOVEABLE | Kernel32.GMEM.GMEM_ZEROINIT);
                handle = hg.DangerousGetHandle();
                if (hg.IsNull) return HRESULT.E_OUTOFMEMORY;

                var lc = Kernel32.GlobalLock(hg);
                StringToAnsiString(s, (byte*)lc.ToPointer(), nb);
                Kernel32.GlobalUnlock(hg);
            }
            return HRESULT.S_OK;
        }

        void IComDataObject.GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            var id = new FormatId(format.cfFormat);
            if (!TryGetData(id, out object val))
            {
                Marshal.ThrowExceptionForHR(unchecked((int)0x80040064));
            }

            switch (val)
            {
                case string s:
                    var ptr = medium.unionmember;
                    SaveStringToHandle(ref ptr, s, true);
                    medium.unionmember = ptr;
                    break;
            }
        }

        int IComDataObject.QueryGetData(ref FORMATETC format)
        {
            throw new NotImplementedException();
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
}
