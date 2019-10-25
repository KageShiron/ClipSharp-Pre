using System;
using System.IO;
using System.Runtime.InteropServices;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace ClipSharp
{
    [Guid("0000000c-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface IStream
    {
        // ISequentialStream portion
        void Read([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] [Out]
            byte[] pv, int cb, out uint pcbRead);

        void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)]
            byte[] pv, int cb, IntPtr pcbWritten);


        // IStream portion
        void Seek(long dlibMove, SeekOrigin dwOrigin, out long plibNewPosition);
        void SetSize(long libNewSize);
        void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten);
        void Commit(int grfCommitFlags);
        void Revert();
        void LockRegion(long libOffset, long cb, int dwLockType);
        void UnlockRegion(long libOffset, long cb, int dwLockType);
        void Stat(out STATSTG pstatstg, STATFLAG grfStatFlag);
        void Clone(out IStream ppstm);
    }

    public enum STATFLAG
    {
        DEFAULT = 0,
        NONAME = 1,
        NOOPEN = 2, //not implemented
    }
}