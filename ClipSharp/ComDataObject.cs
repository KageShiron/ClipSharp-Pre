using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;

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

        public virtual IEnumerable<DataObjectFormat> GetFormats(bool allFormat = false)
        {
            var enumFormatEtc = DataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
            if (enumFormatEtc == null) return null;
            enumFormatEtc.Reset();
            FORMATETC[] fe = new FORMATETC[1];
            List<DataObjectFormat> fs = new List<DataObjectFormat>();
            while(enumFormatEtc.Next(1,fe,null) == 0)
            {
                fs.Add(new DataObjectFormat(fe[0]));
            }
            Marshal.ReleaseComObject(enumFormatEtc);
            return fs;
        }


        public string GetString(FormatId id,NativeStringType type)
        {
            var f = DataObjectUtils.GetFormatEtc(id);
            STGMEDIUM s;
            DataObject.GetData(ref f, out s);
            var str = s.GetString(type);
            s.Release();
            return str;
        }

        public string[] GetFileDropList()
        {
            STGMEDIUM s;
            var f = FormatId.CF_HDROP.FormatEtc;
            DataObject.GetData(ref f, out s);
            var fs = s.GetFiles();
            s.Release();
            return fs;
        }

        public Image GetBitmap()
        {
            var f = FormatId.CF_BITMAP.FormatEtc;
            STGMEDIUM s;
            DataObject.GetData(ref f, out s);
            try
            {
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                var bmp = Bitmap.FromHbitmap(s.unionmember);
                var ret = bmp.Clone() as Bitmap;
                bmp.Dispose();
                return ret;
            }
            finally
            {
                s.Release();
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
