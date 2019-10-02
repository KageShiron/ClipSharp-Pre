using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.CompilerServices;
using System.IO;

namespace ClipSharp
{
    public class ComDataObject
    {

        [DllImport("ole32.dll")]
        static extern int OleGetClipboard(out IDataObject ppDataObj);

        public IDataObject DataObject { get; }

        public ComDataObject()
        {
            IDataObject data;
            OleGetClipboard( out data );
            DataObject = data;
        }

        #region GetCanonicalFormatEtc
        public virtual FORMATETC GetCanonicalFormatEtc( ref FORMATETC format)
        {
            FORMATETC c;
            var re = DataObject.GetCanonicalFormatEtc(ref format, out c);
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
        public virtual FORMATETC GetCanonicalFormatEtc(FormatId format) => GetCanonicalFormatEtc(format.Id);

        #endregion GetCanonicalFormatEtc

        #region GetDataPresent

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(int format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(string format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(FormatId format) => GetDataPresent(format.Id);

        public virtual bool GetDataPresent(ref FORMATETC format)
            => DataObject.QueryGetData(ref format) == 0;//S_OK
        #endregion GetDataPresent

        public virtual IEnumerable<DataObjectFormat> GetFormats(bool allFormat = false)
        {
            IEnumFORMATETC enumFormatEtc = default;
            try
            {
                enumFormatEtc = DataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
                if (enumFormatEtc == null) return null;
                enumFormatEtc.Reset();
                FORMATETC[] fe = new FORMATETC[1];
                List<DataObjectFormat> fs = new List<DataObjectFormat>();
                while (enumFormatEtc.Next(1, fe, null) == 0)
                {
                    fs.Add(new DataObjectFormat(fe[0]));
                }
                return fs;
            }
            finally
            {
                Marshal.ReleaseComObject(enumFormatEtc);
            }
        }


        public string GetString(FormatId id,NativeStringType type)
        {
            var f = DataObjectUtils.GetFormatEtc(id);
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                return s.GetString(type);
            }finally
            {
                s.Dispose();
            }
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

        public Image GetBitmap()
        {
            var f = FormatId.CF_BITMAP.FormatEtc;
            STGMEDIUM s = default;
            try
            {
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                var bmp = Bitmap.FromHbitmap(s.unionmember);
                var ret = bmp.Clone() as Bitmap;
                bmp.Dispose();
                return ret;
            }
            finally
            {
                s.Dispose();
            }
        }

        //public Metafile GetMetafile()
        //{
        //    STGMEDIUM stg = new STGMEDIUM();
        //    var f = FormatId.CF_METAFILEPICT.FormatEtc;
        //    DataObject.GetData(ref f, out stg);
        //    try
        //    {
        //        if (stg.tymed != TYMED.TYMED_MFPICT) throw new ApplicationException();
        //        var hm = new Metafile(stg.unionmember,false);
        //        var ret = (Metafile)hm.Clone();
        //        hm.Dispose();
        //        return ret;
        //    }
        //    finally
        //    {
        //        stg.Release();
        //    }
        //}
        //public Metafile GetEnhancedMetafile()
        //{
        //    STGMEDIUM stg = new STGMEDIUM();
        //    var f = FormatId.CF_ENHMETAFILE.FormatEtc;
        //    DataObject.GetData(ref f, out stg);
        //    try
        //    {
        //        if (stg.tymed != TYMED.TYMED_ENHMF) throw new ApplicationException();
        //        return new Metafile(stg.GetManagedStream());
        //    }
        //    finally
        //    {
        //        stg.Release();
        //    }
        //}

        public string GetString() => GetString(FormatId.CF_UNICODETEXT);

        public string GetString( FormatId id)
        {
            if(id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT || id == FormatId.CommaSeparatedValue )
            {
                return GetString(id, NativeStringType.Ansi);
            }
            else if (id == FormatId.Html || id == FormatId.Xaml)
            {
                return GetString(id, NativeStringType.Utf8);
            }else if (id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust)
            {
                return GetString(id, NativeStringType.Unicode);
            }
            throw new ArgumentException(nameof(id));
        }

    }
}
