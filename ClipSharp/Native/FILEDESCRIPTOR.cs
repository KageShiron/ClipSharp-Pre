using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ClipSharp.Native
{
    [Flags]
    public enum FileDescriptorFlags : uint
    {
        None = 0,
        FD_CLSID = 0x00000001,
        FD_SIZEPOINT = 0x00000002,
        FD_ATTRIBUTES = 0x00000004,
        FD_CREATETIME = 0x00000008,
        FD_ACCESSTIME = 0x00000010,
        FD_WRITESTIME = 0x00000020,
        FD_FILESIZE = 0x00000040,
        FD_PROGRESSUI = 0x00004000,
        FD_LINKUI = 0x00008000,
        FD_UNICODE = 0x80000000 //Windows Vista and later
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public unsafe struct FILEDESCRIPTOR
    {
        public FileDescriptorFlags dwFlags;
        public Guid clsid;
        public SIZE sizel;
        public POINT pointl;
        public FileAttributes dwFileAttributes;
        public long ftCreationTime;
        public long ftLastAccessTime;
        public long ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public fixed char cFileName[260];
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SIZE
    {
        public int cx;
        public int cy;

        public SIZE(int cx, int cy)
        {
            this.cx = cx;
            this.cy = cy;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            X = x;
            Y = y;
        }
    }
}