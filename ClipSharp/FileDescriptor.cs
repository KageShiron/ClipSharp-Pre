using ClipSharp.Native;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ClipSharp
{
    /// <summary>
    /// ファイル記述子を表すクラス
    /// </summary>
    public class FileDescriptor
    {
        private FILEDESCRIPTOR _fd;

        private FileDescriptor(in FILEDESCRIPTOR fd)
        {
            _fd = fd;
            unsafe
            {
                fixed (char* ptr = fd.cFileName)
                {
                    _FileName = new string(ptr);
                }
            }
        }

        /// <summary>
        /// ネイティブのFILEDESCRIPTOR構造体
        /// </summary>
        public FILEDESCRIPTOR FileDesriptor => _fd;


        /// <summary>
        /// FILEGROUPDESCRIPTOR構造体からファイル記述子の配列を作成します
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static FileDescriptor[] FromFileGroupDescriptor(ReadOnlySpan<byte> s)
        {
            var len = MemoryMarshal.Read<uint>(s);
            var fs = MemoryMarshal.Cast<byte, FILEDESCRIPTOR>(s.Slice(sizeof(uint)));
            var list = new FileDescriptor[len];
            for (var i = 0; i < len; i++) list[i] = new FileDescriptor(in fs[i]);
            return list;
        }

        public Guid? Clsid
        {
            get => ValueOrNull(FileDescriptorFlags.FD_CLSID, _fd.clsid);
            set => SetValue(FileDescriptorFlags.FD_CLSID, ref _fd.clsid, value);
        }

        public Rectangle? Rect
        {
            get
            {
                var p = ValueOrNull(FileDescriptorFlags.FD_SIZEPOINT, _fd.pointl);
                return p.HasValue ? (Rectangle?)new Rectangle(p.Value.X, p.Value.Y, _fd.sizel.cx, _fd.sizel.cy) : null;
            }
            set
            {
                if (value.HasValue)
                {
                    _fd.pointl = new POINT(value.Value.X, value.Value.Y);
                    _fd.sizel = new SIZE(value.Value.Width, value.Value.Height);
                    _fd.dwFlags |= FileDescriptorFlags.FD_SIZEPOINT;
                }
                else
                {
                    _fd.dwFlags &= ~FileDescriptorFlags.FD_SIZEPOINT;
                }
            }
        }

        public FileAttributes? FileAttributes
        {
            get => ValueOrNull(FileDescriptorFlags.FD_ATTRIBUTES, _fd.dwFileAttributes);
            set => SetValue(FileDescriptorFlags.FD_ATTRIBUTES, ref _fd.dwFileAttributes, value);
        }

        public DateTime? CreationTime
        {
            get => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_CREATETIME, _fd.ftCreationTime));
            set => SetDateTime(FileDescriptorFlags.FD_CREATETIME, ref _fd.ftCreationTime, value);
        }

        public DateTime? LastAccessTime
        {
            get => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_ACCESSTIME, _fd.ftLastAccessTime));
            set => SetDateTime(FileDescriptorFlags.FD_ACCESSTIME, ref _fd.ftLastAccessTime, value);
        }

        public DateTime? WriteTime
        {
            get => FILETIME2DateTime(ValueOrNull(FileDescriptorFlags.FD_WRITESTIME, _fd.ftLastWriteTime));
            set => SetDateTime(FileDescriptorFlags.FD_WRITESTIME, ref _fd.ftLastWriteTime, value);
        }

        public ulong? FileSize
        {
            get => ValueOrNull(FileDescriptorFlags.FD_FILESIZE, ((ulong)_fd.nFileSizeHigh << 32) | _fd.nFileSizeLow);
            set
            {
                if (value.HasValue)
                {
                    _fd.nFileSizeLow = unchecked((uint)value.Value & 0xFFFFFFFF);
                    _fd.nFileSizeHigh = unchecked((uint)(value.Value >> 32) & 0xFFFFFFFF);
                    _fd.dwFlags |= FileDescriptorFlags.FD_FILESIZE;
                }
                else
                {
                    _fd.dwFlags &= ~FileDescriptorFlags.FD_FILESIZE;
                }
            }
        }

        public string _FileName;

        public string FileName
        {
            get => _FileName;
            set
            {
                if (value.Length >= 260) throw new ArgumentException(value);
                _FileName = value;
                unsafe
                {
                    fixed (char* ptr = _fd.cFileName)
                    {
                        var mem = new Span<char>(ptr, 260);
                        value.AsSpan().CopyTo(mem);
                        mem[value.Length] = '\0';
                    }
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private TResult? ValueOrNull<TResult>(FileDescriptorFlags flag, TResult value) where TResult : struct => _fd.dwFlags.HasFlag(flag) ? value : (TResult?)null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void SetValue<T>(FileDescriptorFlags flag, ref T target, T? value) where T : unmanaged
        {
            if (value.HasValue)
            {
                target = value.Value;
                _fd.dwFlags |= flag;
            }
            else
            {
                _fd.dwFlags &= ~flag;
            }
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void SetDateTime(FileDescriptorFlags flag, ref long target, DateTime? value)
        {
            if (value.HasValue)
            {
                target = value.Value.Ticks;
                _fd.dwFlags |= flag;
            }
            else
            {
                _fd.dwFlags &= ~flag;
            }
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static DateTime? FILETIME2DateTime(long? t)
        {
            if (t.HasValue)
                return DateTime.FromFileTime(t.Value);
            else
                return null;
        }
    }
}