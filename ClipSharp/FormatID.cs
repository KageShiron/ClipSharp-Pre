﻿using System;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ClipSharp
{
    public struct FormatId
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClipboardFormatName(int format
            , [Out] StringBuilder lpszFormatName, int cchMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int RegisterClipboardFormat(string lpszFormat);

        /// <summary>
        /// Create new DataFormatIdentify and add cache.
        /// </summary>
        /// <param name="nativeName"></param>
        /// <returns></returns>
        private static string RegisterFormatId(int id)
        {
            var name = DataObjectUtils.GetFormatName(id);
            _formats.TryAdd(id, new InternalFormatID(name, ""));
            return name;
        }

        /// <summary>
        /// Create new DataFormatIdentify and add cache.
        /// </summary>
        /// <param name="nativeName"></param>
        /// <returns></returns>
        private static FormatId RegisterFormatId(string nativeName)
        {
            var id = DataObjectUtils.GetFormatId(nativeName);
            _formats.TryAdd(id, new InternalFormatID(nativeName, ""));
            return new FormatId(id);
        }

        /// <summary>
        /// Create FormatId from name.
        /// </summary>
        /// <param name="nativeName"></param>
        /// <returns></returns>
        public static FormatId FromNativeName(string nativeName)
        {
            foreach (var (k, v) in _formats)
                if (string.Compare(v.NativeName, nativeName, StringComparison.OrdinalIgnoreCase) == 0)
                    return new FormatId(k);

            return RegisterFormatId(nativeName);
        }

        /// <summary>
        /// Create FormatId from name. (compatible with .NET)
        /// </summary>
        /// <param name="name">format name</param>
        /// <returns>FormatID</returns>
        public static FormatId FromName(string name)
        {
            if (name.StartsWith("Format") && int.TryParse(name.Substring(6), out var num)) return new FormatId(num);
            return name switch
            {
                "Text" => new FormatId((int)CLIPFORMAT.CF_TEXT),
                "UnicodeText" => new FormatId((int)CLIPFORMAT.CF_UNICODETEXT),
                "DeviceIndependentBitmap" => new FormatId((int)CLIPFORMAT.CF_DIB),
                "Bitmap" => new FormatId((int)CLIPFORMAT.CF_BITMAP),
                "EnhancedMetafile" => new FormatId((int)CLIPFORMAT.CF_ENHMETAFILE),
                "MetaFilePict" => new FormatId((int)CLIPFORMAT.CF_METAFILEPICT),
                "SymbolicLink" => new FormatId((int)CLIPFORMAT.CF_SYLK),
                "DataInterchangeFormat" => new FormatId((int)CLIPFORMAT.CF_DIF),
                "TaggedImageFileFormat" => new FormatId((int)CLIPFORMAT.CF_TIFF),
                "OEMText" => new FormatId((int)CLIPFORMAT.CF_OEMTEXT),
                "Palette" => new FormatId((int)CLIPFORMAT.CF_PALETTE),
                "PenData" => new FormatId((int)CLIPFORMAT.CF_PENDATA),
                "RiffAudio" => new FormatId((int)CLIPFORMAT.CF_RIFF),
                "WaveAudio" => new FormatId((int)CLIPFORMAT.CF_WAVE),
                "FileDrop" => new FormatId((int)CLIPFORMAT.CF_HDROP),
                "Locale" => new FormatId((int)CLIPFORMAT.CF_LOCALE),
                _ => FromNativeName(name)
            };
        }

        internal static int GetFormatIdInternal(string formatName)
        {
            var id = RegisterClipboardFormat(formatName);
            if (id == 0) throw new Win32Exception();
            return id;
        }

        static FormatId()
        {
            var fmts = Enum.GetValues(typeof(CLIPFORMAT));
            foreach (CLIPFORMAT c in fmts) _formats.TryAdd((int)c, new InternalFormatID("", c.ToString()));

            var id = GetFormatIdInternal("Shell IDList Array");
            CFSTR_SHELLIDLIST = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Shell IDList Array", "CFSTR_SHELLIDLIST"));
            id = GetFormatIdInternal("Shell Object Offsets");
            CFSTR_SHELLIDLISTOFFSET = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Shell Object Offsets", "CFSTR_SHELLIDLISTOFFSET"));
            id = GetFormatIdInternal("Net Resource");
            CFSTR_NETRESOURCES = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Net Resource", "CFSTR_NETRESOURCES"));
            id = GetFormatIdInternal("FileGroupDescriptor");
            CFSTR_FILEDESCRIPTORA = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileGroupDescriptor", "CFSTR_FILEDESCRIPTORA"));
            id = GetFormatIdInternal("FileGroupDescriptorW");
            CFSTR_FILEDESCRIPTORW = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileGroupDescriptorW", "CFSTR_FILEDESCRIPTORW"));
            id = GetFormatIdInternal("FileContents");
            CFSTR_FILECONTENTS = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileContents", "CFSTR_FILECONTENTS"));
            id = GetFormatIdInternal("FileNameW");
            CFSTR_FILENAMEW = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileNameW", "CFSTR_FILENAMEW"));
            id = GetFormatIdInternal("PrinterFriendlyName");
            CFSTR_PRINTERGROUP = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("PrinterFriendlyName", "CFSTR_PRINTERGROUP"));
            id = GetFormatIdInternal("FileNameMap");
            CFSTR_FILENAMEMAPA = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileNameMap", "CFSTR_FILENAMEMAPA"));
            id = GetFormatIdInternal("FileNameMapW");
            CFSTR_FILENAMEMAPW = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("FileNameMapW", "CFSTR_FILENAMEMAPW"));
            id = GetFormatIdInternal("UniformResourceLocator");
            CFSTR_SHELLURL = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("UniformResourceLocator", "CFSTR_SHELLURL"));
            id = GetFormatIdInternal("UniformResourceLocator");
            CFSTR_INETURLA = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("UniformResourceLocator", "CFSTR_INETURLA"));
            id = GetFormatIdInternal("UniformResourceLocatorW");
            CFSTR_INETURLW = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("UniformResourceLocatorW", "CFSTR_INETURLW"));
            id = GetFormatIdInternal("Preferred DropEffect");
            CFSTR_PREFERREDDROPEFFECT = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Preferred DropEffect", "CFSTR_PREFERREDDROPEFFECT"));
            id = GetFormatIdInternal("Performed DropEffect");
            CFSTR_PERFORMEDDROPEFFECT = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Performed DropEffect", "CFSTR_PERFORMEDDROPEFFECT"));
            id = GetFormatIdInternal("Paste Succeeded");
            CFSTR_PASTESUCCEEDED = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Paste Succeeded", "CFSTR_PASTESUCCEEDED"));
            id = GetFormatIdInternal("InShellDragLoop");
            CFSTR_INDRAGLOOP = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("InShellDragLoop", "CFSTR_INDRAGLOOP"));
            id = GetFormatIdInternal("MountedVolume");
            CFSTR_MOUNTEDVOLUME = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("MountedVolume", "CFSTR_MOUNTEDVOLUME"));
            id = GetFormatIdInternal("PersistedDataObject");
            CFSTR_PERSISTEDDATAOBJECT = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("PersistedDataObject", "CFSTR_PERSISTEDDATAOBJECT"));
            id = GetFormatIdInternal("TargetCLSID");
            CFSTR_TARGETCLSID = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("TargetCLSID", "CFSTR_TARGETCLSID"));
            id = GetFormatIdInternal("Logical Performed DropEffect");
            CFSTR_LOGICALPERFORMEDDROPEFFECT = new FormatId(id);
            _formats.TryAdd(id,
                new InternalFormatID("Logical Performed DropEffect", "CFSTR_LOGICALPERFORMEDDROPEFFECT"));
            id = GetFormatIdInternal("Autoplay Enumerated IDList Array");
            CFSTR_AUTOPLAY_SHELLIDLISTS = new FormatId(id);
            _formats.TryAdd(id,
                new InternalFormatID("Autoplay Enumerated IDList Array", "CFSTR_AUTOPLAY_SHELLIDLISTS"));
            id = GetFormatIdInternal("UntrustedDragDrop");
            CFSTR_UNTRUSTEDDRAGDROP = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("UntrustedDragDrop", "CFSTR_UNTRUSTEDDRAGDROP"));
            id = GetFormatIdInternal("File Attributes Array");
            CFSTR_FILE_ATTRIBUTES_ARRAY = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("File Attributes Array", "CFSTR_FILE_ATTRIBUTES_ARRAY"));
            id = GetFormatIdInternal("InvokeCommand DropParam");
            CFSTR_INVOKECOMMAND_DROPPARAM = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("InvokeCommand DropParam", "CFSTR_INVOKECOMMAND_DROPPARAM"));
            id = GetFormatIdInternal("DropHandlerCLSID");
            CFSTR_SHELLDROPHANDLER = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("DropHandlerCLSID", "CFSTR_SHELLDROPHANDLER"));
            id = GetFormatIdInternal("DropDescription");
            CFSTR_DROPDESCRIPTION = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("DropDescription", "CFSTR_DROPDESCRIPTION"));
            id = GetFormatIdInternal("ZoneIdentifier");
            CFSTR_ZONEIDENTIFIER = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("ZoneIdentifier", "CFSTR_ZONEIDENTIFIER"));
            id = GetFormatIdInternal("Xaml");
            Xaml = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("Xaml", ""));

            id = GetFormatIdInternal("XamlPackage");
            XamlPackage = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("XamlPackage", ""));

            id = GetFormatIdInternal("ApplicationTrust");
            ApplicationTrust = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("ApplicationTrust", ""));

            id = GetFormatIdInternal("HTML Format");
            Html = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("HTML Format", ""));

            id = GetFormatIdInternal("Rich Text Format");
            Rtf = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("XamlPackage", ""));

            id = GetFormatIdInternal("CSV");
            CommaSeparatedValue = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("CSV", ""));

            id = GetFormatIdInternal("PersistentObject");
            Serializable = new FormatId(id);
            _formats.TryAdd(id, new InternalFormatID("PersistentObject", ""));
        }

        private static readonly ConcurrentDictionary<int, InternalFormatID> _formats =
            new ConcurrentDictionary<int, InternalFormatID>();

        private InternalFormatID GetFormatInformation(int id)
        {
            if (_formats.TryGetValue(id, out var val))
            {
                return val;
            }
            else
            {
                RegisterFormatId(id);
                return GetFormatInformation(id);
            }
        }

        public int Id { get; }
        public string NativeName => GetFormatInformation(Id).NativeName;
        public string ConstantName => GetFormatInformation(Id).ConstantName;
        public FORMATETC FormatEtc => DataObjectUtils.GetFormatEtc(this);

        public string DotNetName => (CLIPFORMAT)Id switch
        {
            CLIPFORMAT.CF_TEXT => "Text",
            CLIPFORMAT.CF_UNICODETEXT => "UnicodeText",
            CLIPFORMAT.CF_DIB => "DeviceIndependentBitmap",
            CLIPFORMAT.CF_BITMAP => "Bitmap",
            CLIPFORMAT.CF_ENHMETAFILE => "EnhancedMetafile",
            CLIPFORMAT.CF_METAFILEPICT => "MetaFilePict",
            CLIPFORMAT.CF_SYLK => "SymbolicLink",
            CLIPFORMAT.CF_DIF => "DataInterchangeFormat",
            CLIPFORMAT.CF_TIFF => "TaggedImageFileFormat",
            CLIPFORMAT.CF_OEMTEXT => "OEMText",
            CLIPFORMAT.CF_PALETTE => "Palette",
            CLIPFORMAT.CF_PENDATA => "PenData",
            CLIPFORMAT.CF_RIFF => "RiffAudio",
            CLIPFORMAT.CF_WAVE => "WaveAudio",
            CLIPFORMAT.CF_HDROP => "FileDrop",
            CLIPFORMAT.CF_LOCALE => "Locale",
            _ => NativeName == "" ? "Format" + Id : NativeName
        };

        public FormatId(int id) => Id = id & 0xFFFF;


        public override bool Equals(object obj) => (obj is FormatId) ? Equals((FormatId)obj) : false;

        public bool Equals(FormatId p) => Id == p.Id;

        public static bool operator ==(FormatId lhs, FormatId rhs) => lhs.Equals(rhs);

        public static bool operator !=(FormatId lhs, FormatId rhs) => !lhs.Equals(rhs);

        public override string ToString() => DotNetName;

        public override int GetHashCode() => Id;

        public static readonly FormatId CF_TEXT = new FormatId(1);
        public static readonly FormatId CF_BITMAP = new FormatId(2);
        public static readonly FormatId CF_METAFILEPICT = new FormatId(3);
        public static readonly FormatId CF_SYLK = new FormatId(4);
        public static readonly FormatId CF_DIF = new FormatId(5);
        public static readonly FormatId CF_TIFF = new FormatId(6);
        public static readonly FormatId CF_OEMTEXT = new FormatId(7);
        public static readonly FormatId CF_DIB = new FormatId(8);
        public static readonly FormatId CF_PALETTE = new FormatId(9);
        public static readonly FormatId CF_PENDATA = new FormatId(10);
        public static readonly FormatId CF_RIFF = new FormatId(11);
        public static readonly FormatId CF_WAVE = new FormatId(12);
        public static readonly FormatId CF_UNICODETEXT = new FormatId(13);
        public static readonly FormatId CF_ENHMETAFILE = new FormatId(14);
        public static readonly FormatId CF_HDROP = new FormatId(15);
        public static readonly FormatId CF_LOCALE = new FormatId(16);
        public static readonly FormatId CF_DIBV5 = new FormatId(17);
        public static readonly FormatId CF_OWNERDISPLAY = new FormatId(0x80);
        public static readonly FormatId CF_DSPTEXT = new FormatId(0x81);
        public static readonly FormatId CF_DSPBITMAP = new FormatId(0x82);
        public static readonly FormatId CF_DSPMETAFILEPICT = new FormatId(0x83);
        public static readonly FormatId CF_DSPENHMETAFILE = new FormatId(0x8E);
        public static readonly FormatId CFSTR_SHELLIDLIST;
        public static readonly FormatId CFSTR_SHELLIDLISTOFFSET;
        public static readonly FormatId CFSTR_NETRESOURCES;
        public static readonly FormatId CFSTR_FILEDESCRIPTORA;
        public static readonly FormatId CFSTR_FILEDESCRIPTORW;
        public static readonly FormatId CFSTR_FILECONTENTS;
        public static readonly FormatId CFSTR_FILENAMEA;
        public static readonly FormatId CFSTR_FILENAMEW;
        public static readonly FormatId CFSTR_PRINTERGROUP;
        public static readonly FormatId CFSTR_FILENAMEMAPA;
        public static readonly FormatId CFSTR_FILENAMEMAPW;
        public static readonly FormatId CFSTR_SHELLURL;
        public static readonly FormatId CFSTR_INETURLA;
        public static readonly FormatId CFSTR_INETURLW;
        public static readonly FormatId CFSTR_PREFERREDDROPEFFECT;
        public static readonly FormatId CFSTR_PERFORMEDDROPEFFECT;
        public static readonly FormatId CFSTR_PASTESUCCEEDED;
        public static readonly FormatId CFSTR_INDRAGLOOP;
        public static readonly FormatId CFSTR_MOUNTEDVOLUME;
        public static readonly FormatId CFSTR_PERSISTEDDATAOBJECT;
        public static readonly FormatId CFSTR_TARGETCLSID;
        public static readonly FormatId CFSTR_LOGICALPERFORMEDDROPEFFECT;
        public static readonly FormatId CFSTR_AUTOPLAY_SHELLIDLISTS;
        public static readonly FormatId CFSTR_UNTRUSTEDDRAGDROP;
        public static readonly FormatId CFSTR_FILE_ATTRIBUTES_ARRAY;
        public static readonly FormatId CFSTR_INVOKECOMMAND_DROPPARAM;
        public static readonly FormatId CFSTR_SHELLDROPHANDLER;
        public static readonly FormatId CFSTR_DROPDESCRIPTION;
        public static readonly FormatId CFSTR_ZONEIDENTIFIER;
        public static readonly FormatId Html;
        public static readonly FormatId Rtf;
        public static readonly FormatId CommaSeparatedValue;
        public static readonly FormatId Serializable;
        public static readonly FormatId Xaml;
        public static readonly FormatId XamlPackage;
        public static readonly FormatId ApplicationTrust;
    }

    internal class InternalFormatID
    {
        public InternalFormatID(string nativeName, string constantName)
        {
            NativeName = nativeName;
            ConstantName = constantName;
        }

        public string NativeName { get; set; }
        public string ConstantName { get; set; }
    }
}