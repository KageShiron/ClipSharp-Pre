using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ClipSharp
{
    public class DataObjectFormat
    {
        private Dictionary<string, string> cfstrs = new Dictionary<string, string>()
        {
            {"Shell IDList Array", "CFSTR_SHELLIDLIST"}, // CF_IDLIST
            {"Shell Object Offsets", "CFSTR_SHELLIDLISTOFFSET"}, // CF_OBJECTPOSITIONS
            {"Net Resource", "CFSTR_NETRESOURCES"}, // CF_NETRESOURCE
            {"FileGroupDescriptor", "CFSTR_FILEDESCRIPTORA"}, // CF_FILEGROUPDESCRIPTORA
            {"FileGroupDescriptorW", "CFSTR_FILEDESCRIPTORW"}, // CF_FILEGROUPDESCRIPTORW
            {"FileContents", "CFSTR_FILECONTENTS"}, // CF_FILECONTENTS
            {"FileName", "CFSTR_FILENAMEA"}, // CF_FILENAMEA
            {"FileNameW", "CFSTR_FILENAMEW"}, // CF_FILENAMEW
            {"PrinterFriendlyName", "CFSTR_PRINTERGROUP"}, // CF_PRINTERS
            {"FileNameMap", "CFSTR_FILENAMEMAPA"}, // CF_FILENAMEMAPA
            {"FileNameMapW", "CFSTR_FILENAMEMAPW"}, // CF_FILENAMEMAPW
            {"UniformResourceLocator", "CFSTR_SHELLURL"},
            //{"", "CFSTR_INETURLA"},
            {"UniformResourceLocatorW", "CFSTR_INETURLW"},
            {"Preferred DropEffect", "CFSTR_PREFERREDDROPEFFECT"},
            {"Performed DropEffect", "CFSTR_PERFORMEDDROPEFFECT"},
            {"Paste Succeeded", "CFSTR_PASTESUCCEEDED"},
            {"InShellDragLoop", "CFSTR_INDRAGLOOP"},
            {"MountedVolume", "CFSTR_MOUNTEDVOLUME"},
            {"PersistedDataObject", "CFSTR_PERSISTEDDATAOBJECT"},
            {"TargetCLSID", "CFSTR_TARGETCLSID"}, // HGLOBAL with a CLSID of the drop target
            {"Logical Performed DropEffect", "CFSTR_LOGICALPERFORMEDDROPEFFECT"},
            {"Autoplay Enumerated IDList Array", "CFSTR_AUTOPLAY_SHELLIDLISTS"}, // (HGLOBAL with LPIDA,
            {"UntrustedDragDrop", "CFSTR_UNTRUSTEDDRAGDROP"}, //  DWORD
            {"File Attributes Array", "CFSTR_FILE_ATTRIBUTES_ARRAY"}, // (FILE_ATTRIBUTES_ARRAY format on HGLOBAL,
            {"InvokeCommand DropParam", "CFSTR_INVOKECOMMAND_DROPPARAM"}, // (HGLOBAL with LPWSTR,
            {"DropHandlerCLSID", "CFSTR_SHELLDROPHANDLER"}, // (HGLOBAL with CLSID of drop handler,
            {"DropDescription", "CFSTR_DROPDESCRIPTION"}, // (HGLOBAL with DROPDESCRIPTION,
            {"ZoneIdentifier", "CFSTR_ZONEIDENTIFIER"}, //  DWORD, to be used with CFSTR_FILECONTENTS data transfers
        };

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

        public Exception Error { get; }
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

