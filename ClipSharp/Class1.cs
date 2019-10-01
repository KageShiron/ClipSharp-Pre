using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ClipSharp
{
    public class FileDescriptor
    {
        private FILEDESCRIPTOR _fd;
        private FileDescriptor(in FILEDESCRIPTOR fd)
        {
            this._fd = fd;
        }

        public static FileDescriptor[] FromFileGroupDescriptor(ReadOnlySpan<byte> s)
        {
            var len = MemoryMarshal.Read<uint>(s);
            var fs = MemoryMarshal.Cast<byte, FILEDESCRIPTOR>(s.Slice(sizeof(uint)));
            var list = new FileDescriptor[len];
            for (int i = 0; i < len; i++)
            {
                list[i] = new FileDescriptor(in fs[i]);
            }
            return list;
        }

        public Guid? Clsid => ValueOrNull(FileDescriptorFlags.FD_CLSID, _fd.clsid);
        public SIZE? Size => ValueOrNull(FileDescriptorFlags.FD_SIZEPOINT, _fd.sizel);
        public POINT? Point => ValueOrNull(FileDescriptorFlags.FD_SIZEPOINT, _fd.pointl);
        public FileAttributes? FileAttributes => ValueOrNull(FileDescriptorFlags.FD_ATTRIBUTES, _fd.dwFileAttributes);
        public DateTime? CreationTime => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_CREATETIME, _fd.ftCreationTime));
        public DateTime? LastAccessTime => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_ACCESSTIME, _fd.ftLastAccessTime));
        public DateTime? WriteTime => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_WRITESTIME, _fd.ftLastWriteTime));
        public ulong? FileSize => ValueOrNull(FileDescriptorFlags.FD_FILESIZE, (((ulong)_fd.nFileSizeHigh) << 32 | _fd.nFileSizeLow));
        public string FileName => _fd.cFileName;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private TResult? ValueOrNull<TResult>(FileDescriptorFlags flag, TResult value) where TResult : struct
        {
            if ((_fd.dwFlags & flag) == flag) return value; else return null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        static DateTime? FILETIME2DateTime(long? t)
        {
            if (t.HasValue)
                return DateTime.FromFileTime(t.Value);
            else
                return null;
        }

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
        private struct FILEDESCRIPTOR
        {
            public FileDescriptorFlags dwFlags;
            public Guid clsid;
            public SIZE sizel;
            public POINT pointl;
            public FileAttributes dwFileAttributes;
            public long ftCreationTime;
            public long ftLastAccessTime;
            public long ftLastWriteTime;
            public UInt32 nFileSizeHigh;
            public UInt32 nFileSizeLow;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
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
                this.X = x;
                this.Y = y;
            }

        }
    }
}