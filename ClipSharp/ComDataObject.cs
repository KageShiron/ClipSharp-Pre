using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Vanara.PInvoke.Shell32;
using System.Linq;

namespace ClipSharp
{
    public class ComDataObject : System.Windows.Forms.IDataObject
    {
        public ComDataObject(IDataObject data)
        {
            DataObject = data;
        }

        public IDataObject DataObject { get; }

        public virtual IEnumerable<DataObjectFormat> GetFormats(bool allFormat = false)
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


        public HtmlFormat GetHtml()
        {
            return HtmlFormat.Parse(GetString(FormatId.Html));
        }

        public bool GetHtml( out HtmlFormat result)
        {
            if (GetDataPresent(FormatId.Html))
            {
                result = GetHtml();
                return true;
            }
            result = null;
            return false;
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

        public bool GetFileDropList( out string[] result)
        {
            if (GetDataPresent(FormatId.CF_HDROP))
            {
                result = GetFileDropList();
                return true;
            }
            result = null;
            return false;
        }

        public Image GetBitmap()
        {
            var f = FormatId.CF_BITMAP.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                var bmp = Image.FromHbitmap(s.unionmember);
                var ret = (Image) bmp.Clone();
                bmp.Dispose();
                return ret;
            }
            finally
            {
                s.Dispose();
            }
        }
        
        public bool GetBitmap( out Image result)
        {
            if (GetDataPresent(FormatId.CF_BITMAP))
            {
                result = GetBitmap();
                return true;
            }
            result = null;
            return false;
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

        public DragDropEffects GetDragDropEffects()
        {
            return ReadHGlobal<DragDropEffects>(FormatId.CFSTR_PREFERREDDROPEFFECT);
        }

        public bool GetDragDropEffects( out DragDropEffects result)
        {
            if (GetDataPresent(FormatId.CFSTR_PREFERREDDROPEFFECT))
            {
                result = GetDragDropEffects();
                return true;
            }
            result = default;
            return false;
        }

        public CultureInfo GetCultureInfo()
        {
            return new CultureInfo(ReadHGlobal<int>(FormatId.CF_LOCALE));
        }
        public bool GetCultureInfo( out CultureInfo result)
        {
            if (GetDataPresent(FormatId.CF_LOCALE))
            {
                result = GetCultureInfo();
                return true;
            }
            result = null;
            return false;
        }

        public List<PIDL> GetCida()
        {
            var f = FormatId.CFSTR_SHELLIDLIST.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);

                return s.InvokeHGlobal<uint, List<PIDL>>(x =>
                {
                    var l = new List<PIDL>();
                    unsafe
                    {
                        fixed (void* ptr = x)
                        {
                            for (var i = 0; i < x[0] + 1; i++) l.Add(new PIDL((IntPtr) ((byte*) ptr + x[i + 1]), true));
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

        public bool GetCida( out List<PIDL> result)
        {
            if (GetDataPresent(FormatId.CFSTR_SHELLIDLIST))
            {
                result = GetCida();
                return true;
            }
            result = null;
            return false;
        }

        public (List<PIDL>,PIDL) GetShellIdList()
        {
            var f = FormatId.CFSTR_SHELLIDLIST.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);

                return s.InvokeHGlobal<uint, (List<PIDL>,PIDL)>(x =>
                {
                    var l = new List<PIDL>();
                    PIDL? parent = null;
                    unsafe
                    {
                        fixed (void* ptr = x)
                        {
                            parent = new PIDL((IntPtr)((byte*)ptr + x[1]), true, true);
                            for (var i = 1; i < x[0] + 1; i++)
                            {
                                // DO NOT release memory of child
                                var child = new PIDL((IntPtr)((byte*)ptr + x[i + 1]), false, false);
                                var p = new PIDL(parent);
                                p.Append(child);
                                l.Add(p);
                            }
                        }
                    }

                    return (l,parent);
                });
            }
            finally
            {
                s.Dispose();
            }

        }

        public bool GetShellIdList( out List<PIDL>? result , out PIDL? parent )
        {
            if (GetDataPresent(FormatId.CFSTR_SHELLIDLIST))
            {
                (result,parent) = GetShellIdList();
                return true;
            }
            result = null;
            parent = null;
            return false;
        }


        public Stream GetFileContent(int index)
        {
            return GetUnmanagedStream(FormatId.CFSTR_FILECONTENTS, index);
        }
        public bool GetFileContent( int index, out Stream result)
        {
            if (GetDataPresent(FormatId.CFSTR_FILECONTENTS))
            {
                result = GetFileContent(index);
                return true;
            }
            result = null;
            return false;
        }

        public Dictionary<FileDescriptor, Stream> GetFileContents()
        {
            var fd = GetFileDescriptors();
            var s = new Dictionary<FileDescriptor, Stream>(fd.Length);
            for (var i = 0; i < fd.Length; i++) s.Add(fd[i], GetFileContent(i));
            return s;
        }

        public bool GetFileContents( out Dictionary<FileDescriptor, Stream> result)
        {
            if (GetDataPresent(FormatId.CFSTR_FILEDESCRIPTORW))
            {
                result = GetFileContents();
                return true;
            }
            result = null;
            return false;
        }
        public Metafile GetMetafile()
        {
            var f = FormatId.CF_METAFILEPICT.FormatEtc;
            DataObject.GetData(ref f, out var stg);
            try
            {
                if (stg.tymed != TYMED.TYMED_MFPICT) throw new ApplicationException();
                var hm = new Metafile(stg.unionmember, false);
                var ret = (Metafile) hm.Clone();
                hm.Dispose();
                return ret;
            }
            finally
            {
                stg.Dispose();
            }
        }
        public bool GetMetafile( out Metafile result)
        {
            if (GetDataPresent(FormatId.CF_METAFILEPICT))
            {
                result = GetMetafile();
                return true;
            }
            result = null;
            return false;
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

        public bool GetEnhancedMetafile( out Metafile result)
        {
            if (GetDataPresent(FormatId.CF_ENHMETAFILE))
            {
                result = GetEnhancedMetafile();
                return true;
            }
            result = null;
            return false;
        }

        #region GetString

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

        public string GetString()
        {
            return GetString(FormatId.CF_UNICODETEXT);
        }

        public string GetString(FormatId id)
        {
            if (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                id == FormatId.CommaSeparatedValue)
                return GetString(id, NativeStringType.Ansi);
            if (id == FormatId.Html || id == FormatId.Xaml)
                return GetString(id, NativeStringType.Utf8);
            if (id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust)
                return GetString(id, NativeStringType.Unicode);
            throw new ArgumentException(nameof(id));
        }


        public bool GetString( out string result)
        {
            if (GetDataPresent(FormatId.CF_UNICODETEXT))
            {
                result = GetString();
                return true;
            }
            result = null;
            return false;
        }

        public bool GetString(FormatId id, out string result)
        {
            if (GetDataPresent(id))
            {
                result = GetString(id);
                return true;
            }
            result = null;
            return false;
        }


        public bool GetString(FormatId id , NativeStringType native, out string result)
        {
            if (GetDataPresent(id))
            {
                result = GetString(id,native);
                return true;
            }
            result = null;
            return false;
        }
#endregion GetSring


        public Stream GetUnmanagedStream(FormatId id, int lindex = -1)
        {
            var f = id.FormatEtc;
            f.lindex = lindex;
            DataObject.GetData(ref f, out var s);
            return s.GetUnmanagedStream(true);
        }
        public bool GetUnmanagedStream(FormatId id, out Stream result, int lindex = -1)
        {
            if (GetDataPresent(id))
            {
                result = GetUnmanagedStream(id, lindex);
                return true;
            }
            result = null;
            return false;
        }

        public bool GetStream(FormatId id, out Stream result, int lindex = -1)
        {
            if (GetDataPresent(id))
            {
                result = GetStream(id, lindex);
                return true;
            }
            result = null;
            return false;
        }

        public Stream GetStream(FormatId id, int lindex = -1)
        {
            var f = id.FormatEtc;
            f.lindex = lindex;
            DataObject.GetData(ref f, out var s);
            return s.GetManagedStream();
        }

        public FileDescriptor[] GetFileDescriptors()
        {
            var f = FormatId.CFSTR_FILEDESCRIPTORW.FormatEtc;
            DataObject.GetData(ref f, out var s);
            return s.InvokeHGlobal<byte, FileDescriptor[]>(FileDescriptor.FromFileGroupDescriptor);
        }
        public bool GetFileDescriptors( out FileDescriptor[] result)
        {
            var f = FormatId.CFSTR_FILEDESCRIPTORW;
            if (this.GetDataPresent(f))
            {
                result = GetFileDescriptors();
                return true;
            }
            result = null;
            return false;
        }



        #region GetFormats
        string[] System.Windows.Forms.IDataObject.GetFormats()
        {
            return ((System.Windows.Forms.IDataObject)this).GetFormats(true);
        }

        string[] System.Windows.Forms.IDataObject.GetFormats(bool autoConvert)
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

        #region GetCanonicalFormatEtc

        public virtual FORMATETC GetCanonicalFormatEtc(ref FORMATETC format)
        {
            DataObject.GetCanonicalFormatEtc(ref format, out var c);
            return c;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public FORMATETC GetCanonicalFormatEtc(int format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetCanonicalFormatEtc(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public FORMATETC GetCanonicalFormatEtc(string format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetCanonicalFormatEtc(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public virtual FORMATETC GetCanonicalFormatEtc(FormatId format)
        {
            return GetCanonicalFormatEtc(format.Id);
        }

        #endregion GetCanonicalFormatEtc
    }
}