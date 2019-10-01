using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ClipSharp
{
    public static class StgMediumExtensions
    {

        #region P/Invoke

        [DllImport("ole32.dll")]
        static extern void ReleaseStgMedium([In] ref STGMEDIUM pmedium);

        [DllImport("kernel32.dll")]
        private static extern UIntPtr GlobalSize(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern unsafe bool GlobalUnlock(void* hMem);

        [DllImport("gdi32.dll")]
        private static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, [Out] byte[] lpbBuffer);

        [DllImport("gdi32.dll")]
        private static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, IntPtr lpbBuffer);

        [DllImport("gdi32.dll")]
        private static extern uint GetMetaFileBitsEx(IntPtr hmf, uint cbBuffer, [Out] byte[] lpbBuffer);

        [DllImport("gdi32.dll")]
        private static extern uint GetMetaFileBitsEx(IntPtr hmf, uint cbBuffer, IntPtr lpbBuffer);

        [DllImport("ole32.dll")]
        static extern int CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out IStream ppstm);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern unsafe int DragQueryFile(IntPtr hDrop, int iFile,
            char* lpszFile, uint cch);


        #endregion



        public static string GetString(this in STGMEDIUM stg, NativeStringType type = NativeStringType.Unicode)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL) throw new ArgumentException(nameof(stg));

            IntPtr ptr = GlobalLock(stg.unionmember);
            var str = type switch
            {
                NativeStringType.Unicode => Marshal.PtrToStringUni(ptr),
                NativeStringType.Ansi => Marshal.PtrToStringAnsi(ptr),
                NativeStringType.Utf8 => Marshal.PtrToStringUTF8(ptr),
                _ => throw new InvalidCastException()
            };
            GlobalUnlock(ptr);
            return str;
        }
        public unsafe static string[] GetFiles(this STGMEDIUM stg)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL) throw new ArgumentException("tymed is not hglobal");

            int count = DragQueryFile(stg.unionmember, -1, null, 0);
            if (count <= 0) return Array.Empty<string>();
            string[] files = new string[count];

            char* sb = stackalloc char[260];
            for (int i = 0; i < count; i++)
            {
                if (DragQueryFile(stg.unionmember, i, sb, 260) > 0)
                {
                    files[i] = new string(sb);
                }
            }
            return files;
        }
        /// <summary>
        /// STGMEDIUMを解放します
        /// </summary>
        /// <param name="stg">解放するSTGMEDIUM</param>
        public static void Release(this STGMEDIUM stg)
        {
            ReleaseStgMedium(ref stg);
        }


        private static Span<byte> BeginHGlobal(IntPtr hGlobal)
        {
            IntPtr locked = GlobalLock(hGlobal);
            int size = (int)GlobalSize(locked).ToUInt32();
            unsafe
            {
                return new Span<byte>((byte*)locked, size);
            }
        }

        private static void EndHGlobal<T>(Span<T> hGlobal) where T : unmanaged
        {
            unsafe
            {
                fixed (void* ptr = hGlobal)
                {
                    GlobalUnlock(ptr);

                }
            }
        }

    }
    public enum NativeStringType
    {
        Unicode,
        Ansi,
        Utf8
    }
}
