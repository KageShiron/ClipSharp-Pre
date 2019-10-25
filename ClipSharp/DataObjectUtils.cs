using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ClipSharp
{
    public static class DataObjectUtils
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern unsafe int GetClipboardFormatName(int format
            , char* lpszFormatName, int cchMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int RegisterClipboardFormat(string lpszFormat);

        public static unsafe string GetFormatName(int formatId)
        {
            var sb = stackalloc char[260];
            if (GetClipboardFormatName(formatId, sb, 260) == 0) return ""; //$"Format{formatId}";
            return new string(sb);
        }

        public static int GetFormatId(string formatName)
        {
            //if (formatName.StartsWith("Format")) return int.Parse(formatName.Substring(6));
            var id = RegisterClipboardFormat(formatName);
            if (id == 0) throw new Win32Exception();
            return id;
        }


        public static FORMATETC GetFormatEtc(short id, int lindex = -1, DVASPECT dwAspect = DVASPECT.DVASPECT_CONTENT)
        {
            return new FORMATETC
            {
                cfFormat = id,
                dwAspect = dwAspect,
                lindex = lindex,
                tymed = TYMED.TYMED_HGLOBAL | TYMED.TYMED_GDI | TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE |
                        TYMED.TYMED_GDI | TYMED.TYMED_FILE | TYMED.TYMED_MFPICT | TYMED.TYMED_ENHMF
            };
        }

        public static FORMATETC GetFormatEtc(string dataFormat, int lindex = -1,
            DVASPECT dwAspect = DVASPECT.DVASPECT_CONTENT)
        {
            return GetFormatEtc((short)GetFormatId(dataFormat), lindex, dwAspect);
        }

        public static FORMATETC GetFormatEtc(int id, int lindex = -1, DVASPECT dwAspect = DVASPECT.DVASPECT_CONTENT)
        {
            return GetFormatEtc((short)id, lindex, dwAspect);
        }

        public static FORMATETC GetFormatEtc(FormatId id, int lindex = -1,
            DVASPECT dwAspect = DVASPECT.DVASPECT_CONTENT)
        {
            return GetFormatEtc((short)id.Id, lindex, dwAspect);
        }
    }


    public enum CLIPFORMAT
    {
        CF_TEXT = 1,
        CF_BITMAP = 2,
        CF_METAFILEPICT = 3,
        CF_SYLK = 4,
        CF_DIF = 5,
        CF_TIFF = 6,
        CF_OEMTEXT = 7,
        CF_DIB = 8,
        CF_PALETTE = 9,
        CF_PENDATA = 10,
        CF_RIFF = 11,
        CF_WAVE = 12,
        CF_UNICODETEXT = 13,
        CF_ENHMETAFILE = 14,
        CF_HDROP = 15,
        CF_LOCALE = 16,
        CF_DIBV5 = 17,
        CF_OWNERDISPLAY = 0x80,
        CF_DSPTEXT = 0x81,
        CF_DSPBITMAP = 0x82,
        CF_DSPMETAFILEPICT = 0x83,
        CF_DSPENHMETAFILE = 0x8E
    }
}