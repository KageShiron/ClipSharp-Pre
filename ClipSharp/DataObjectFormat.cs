using System;
using System.Runtime.InteropServices.ComTypes;

namespace ClipSharp
{
    public class DataObjectFormat
    {
        public DataObjectFormat(FORMATETC f, int? cannonical = null, bool notDataObject = false)
        {
            FormatId = new FormatId(f.cfFormat);
            if (notDataObject)
            {
                NotDataObject = true;
                return;
            }

            DvAspect = f.dwAspect;
            PtdNull = f.ptd;
            LIndex = f.lindex;
            Tymed = f.tymed;
            Canonical = cannonical; // man.GetCanonicalFormatEtc(f.cfFormat).cfFormat;
        }

        public Exception? Error { get; }
        public FormatId FormatId { get; }
        public DVASPECT DvAspect { get; }
        public IntPtr PtdNull { get; }
        public int LIndex { get; }
        public TYMED Tymed { get; }
        public int? Canonical { get; }
        public bool NotDataObject { get; }

        public override string ToString()
        {
            return FormatId.NativeName;
        }
    }
}