using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Vanara.PInvoke;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDataObject = System.Windows.Forms.IDataObject;

namespace ClipSharp
{
    public class OleDataObject : IDataObject
    {
        public OleDataObject(IComDataObject data)
        {
            DataObject = data;
        }

        public IComDataObject DataObject { get; }

        public string GetString(FormatId id, NativeStringType type)
        {
            var f = DataObjectUtils.GetFormatEtc(id);
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                return s.GetString(type);
            }
            finally
            {
                s.Dispose();
            }
        }

        public string GetString(FormatId id)
        {
            if (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                id == FormatId.CommaSeparatedValue || id == FormatId.CFSTR_INETURLA)
                return GetString(id, NativeStringType.Ansi);
            if (id == FormatId.Html || id == FormatId.Xaml)
                return GetString(id, NativeStringType.Utf8);
            if (id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust || id == FormatId.CFSTR_INETURLW)
                return GetString(id, NativeStringType.Unicode);
            throw new ArgumentException(nameof(id));
        }

        public string GetString()
        {
            return GetString(FormatId.CF_UNICODETEXT);
        }

        public HtmlFormat GetHtml()
        {
            return HtmlFormat.Parse(GetString(FormatId.Html));
        }

        public string[] GetFileDropList()
        {
            var f = FormatId.CF_HDROP.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                return s.GetFiles();
            }
            finally
            {
                s.Dispose();
            }
        }

        public DragDropEffects GetDragDropEffects()
        {
            return ReadHGlobal<DragDropEffects>(FormatId.CFSTR_PREFERREDDROPEFFECT);
        }

        public CultureInfo GetCultureInfo()
        {
            return new CultureInfo(ReadHGlobal<int>(FormatId.CF_LOCALE));
        }

        public List<Shell32.PIDL> GetPidl()
        {
            var f = FormatId.CFSTR_SHELLIDLIST.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);

                return s.InvokeHGlobal<uint, List<Shell32.PIDL>>((_, x) =>
                {
                    var l = new List<Shell32.PIDL>();
                    unsafe
                    {
                        fixed (void* ptr = x)
                        {
                            for (var i = 0; i < x[0] + 1; i++)
                                l.Add(new Shell32.PIDL((IntPtr)((byte*)ptr + x[i + 1]), true));
                        }
                    }

                    return l;
                });
            }
            finally
            {
                s.Dispose();
            }
        }

        public FileDescriptor[] GetFileDescriptors()
        {
            var f = FormatId.CFSTR_FILEDESCRIPTORW.FormatEtc;
            DataObject.GetData(ref f, out var s);
            return s.InvokeHGlobal<byte, FileDescriptor[]>((_, f) => FileDescriptor.FromFileGroupDescriptor(f));
        }

        public Stream GetFileContent(int index)
        {
            return GetUnmanagedStream(FormatId.CFSTR_FILECONTENTS, index);
        }

        public Dictionary<FileDescriptor, Stream> GetFileContents()
        {
            var fd = GetFileDescriptors();
            var s = new Dictionary<FileDescriptor, Stream>(fd.Length);
            for (var i = 0; i < fd.Length; i++) s.Add(fd[i], GetFileContent(i));
            return s;
        }

        public Image GetBitmap(FormatId id)
        {
            if (id == FormatId.CF_BITMAP) return GetBitmap();
            var f = id.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                using var st = s.GetUnmanagedStream(true);
                using var bmp = Image.FromStream(st);
                var ret = (Image)bmp.Clone();
                return ret;
            }
            finally
            {
                s.Dispose();
            }
        }

        public Image GetBitmap()
        {
            var f = FormatId.CF_BITMAP.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                using var bmp = Image.FromHbitmap(s.unionmember);
                var ret = (Image)bmp.Clone();
                return ret;
            }
            finally
            {
                s.Dispose();
            }
        }

        public Metafile GetMetafile()
        {
            var f = FormatId.CF_METAFILEPICT.FormatEtc;
            DataObject.GetData(ref f, out var stg);
            try
            {
                if (stg.tymed != TYMED.TYMED_MFPICT) throw new ApplicationException();
                using var hm = new Metafile(stg.unionmember, false);
                var ret = (Metafile)hm.Clone();
                return ret;
            }
            finally
            {
                stg.Dispose();
            }
        }

        public Metafile GetEnhancedMetafile()
        {
            var f = FormatId.CF_ENHMETAFILE.FormatEtc;
            DataObject.GetData(ref f, out var stg);
            try
            {
                if (stg.tymed != TYMED.TYMED_ENHMF) throw new ApplicationException();
                return new Metafile(stg.GetManagedStream());
            }
            finally
            {
                stg.Dispose();
            }
        }

        public Stream GetStream(FormatId id, int lindex = -1)
        {
            var f = id.FormatEtc;
            f.lindex = lindex;
            DataObject.GetData(ref f, out var s);
            return s.GetManagedStream();
        }

        public TResult ReadHGlobal<TResult>(FormatId id) where TResult : unmanaged
        {
            STGMEDIUM s = default;
            var f = id.FormatEtc;
            try
            {
                DataObject.GetData(ref f, out s);
                return s.ReadHGlobal<TResult>();
            }
            finally
            {
                s.Dispose();
            }
        }

        private Stream GetUnmanagedStream(FormatId id, int lindex = -1)
        {
            var f = id.FormatEtc;
            f.lindex = lindex;
            DataObject.GetData(ref f, out var s);
            return s.GetUnmanagedStream(true);
        }

        #region GetFormats
        string[] IDataObject.GetFormats()
        {
            return ((IDataObject)this).GetFormats(true);
        }

        string[] IDataObject.GetFormats(bool autoConvert)
        {
            return GetFormats().Select(x => x.ToString()).ToArray();
        }

        public IEnumerable<DataObjectFormat> GetFormats()
        {
            IEnumFORMATETC enumFormatEtc = null!;
            try
            {
                enumFormatEtc = DataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
                if (enumFormatEtc == null) return Array.Empty<DataObjectFormat>();
                enumFormatEtc.Reset();
                var fe = new FORMATETC[1];
                var fs = new List<DataObjectFormat>();
                while (enumFormatEtc.Next(1, fe, null) == 0) fs.Add(new DataObjectFormat(fe[0]));
                return fs;
            }
            finally
            {
                if (enumFormatEtc != null)
                    Marshal.ReleaseComObject(enumFormatEtc);
            }
        }
        #endregion

        #region GetDataPresent

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(int format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        public bool GetDataPresent(string format, bool autoConvert)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(string format)
        {
            return GetDataPresent(format, true);
        }

        public bool GetDataPresent(Type format)
        {
            return GetDataPresent(format.FullName);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(FormatId format)
        {
            return GetDataPresent(format.Id);
        }

        public bool GetDataPresent(ref FORMATETC format)
        {
            return DataObject.QueryGetData(ref format) == 0;
        }

        #endregion GetDataPresent

        #region GetData

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object GetData(int format)
        {
            return GetData(new FormatId(format));
        }

        public object GetData(string format, bool autoConvert)
        {
            return GetData(FormatId.FromName(format));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object GetData(string format)
        {
            return GetData(format, true);
        }

        public object GetData(Type format)
        {
            return GetData(format.FullName);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object GetData(FormatId id)
        {
            if (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                id == FormatId.CommaSeparatedValue || id == FormatId.Html || id == FormatId.Xaml ||
                id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust)
                return GetString(id);
            if (id == FormatId.CF_HDROP) return GetFileDropList();
            if (id == FormatId.CFSTR_FILENAMEW)
                return new[] { GetString(FormatId.CFSTR_FILENAMEW) };
            if (id == FormatId.CFSTR_FILENAMEA)
                return new[] { GetString(FormatId.CFSTR_FILENAMEA) };
            if (id == FormatId.CF_BITMAP) return GetBitmap();
            if (id == FormatId.CF_ENHMETAFILE) return GetEnhancedMetafile();
            if (id == FormatId.CF_METAFILEPICT) return GetMetafile();
            return GetStream(id);
        }

        #endregion

        #region SetData

        public void SetData(string format, bool autoConvert, object data) => throw new NotImplementedException();

        public void SetData(string format, object data) => SetData(format, true, data);

        public void SetData(Type format, object data) => SetData(format.FullName, data);

        public void SetData(object data) => throw new NotImplementedException();
        #endregion
    }
}