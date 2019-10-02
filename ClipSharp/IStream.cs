using System;
using System.IO;
using System.Runtime.InteropServices;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace DataFormatLib
{
    [Guid("0000000c-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface IStream
    {
        // ISequentialStream portion
        void Read([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1), Out] Byte[] pv, int cb, out uint pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] Byte[] pv, int cb, IntPtr pcbWritten);

        // IStream portion
        void Seek(Int64 dlibMove, SeekOrigin dwOrigin, out Int64 plibNewPosition);
        void SetSize(Int64 libNewSize);
        void CopyTo(IStream pstm, Int64 cb, IntPtr pcbRead, IntPtr pcbWritten);
        void Commit(int grfCommitFlags);
        void Revert();
        void LockRegion(Int64 libOffset, Int64 cb, int dwLockType);
        void UnlockRegion(Int64 libOffset, Int64 cb, int dwLockType);
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
