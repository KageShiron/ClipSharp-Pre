using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ClipSharp
{
    public delegate TResult ReadOnlySpanFunc<T, TResult>(IntPtr ptr, ReadOnlySpan<T> span);

    public static class StgMediumExtensions
    {
        #region P/Invoke

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(in STGMEDIUM pmedium);

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
        private static extern HRESULT CreateStreamOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out IStream ppstm);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern unsafe int DragQueryFile(IntPtr hDrop, int iFile,
            char* lpszFile, uint cch);

        #endregion


        /// <summary>
        /// Ansi、Unicode、UTF8をSystem.Stringに変換します
        /// </summary>
        /// <param name="stg"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetString(this in STGMEDIUM stg, NativeStringType type = NativeStringType.Unicode)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL) throw new ArgumentException(nameof(stg));

            var ptr = GlobalLock(stg.unionmember);
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

        public static unsafe string[] GetFiles(this STGMEDIUM stg)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL) throw new ArgumentException("tymed is not hglobal");

            var count = DragQueryFile(stg.unionmember, -1, null, 0);
            if (count <= 0) return Array.Empty<string>();
            var files = new string[count];

            var sb = stackalloc char[260];
            for (var i = 0; i < count; i++)
                if (DragQueryFile(stg.unionmember, i, sb, 260) > 0)
                    files[i] = new string(sb);
            return files;
        }

        /// <summary>
        /// STGMEDIUMを解放します
        /// </summary>
        /// <param name="stg">解放するSTGMEDIUM</param>
        public static void Dispose(this in STGMEDIUM stg)
        {
            if (stg.tymed == TYMED.TYMED_NULL) return;
            ReleaseStgMedium(in stg);
        }


        /// <summary>
        /// HGLOBALについてfuncに与えられた処理を実行し、戻り値を返します。
        /// </summary>
        /// <typeparam name="TSpan"></typeparam>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="stg"></param>
        /// <param name="func"></param>
        /// <returns></returns>
        public static TResult InvokeHGlobal<TSpan, TResult>(this in STGMEDIUM stg,
            ReadOnlySpanFunc<TSpan, TResult> func)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL && stg.tymed != TYMED.TYMED_MFPICT) throw new ArgumentException(nameof(stg));
            var locked = GlobalLock(stg.unionmember);
            try
            {
                unsafe
                {
                    var size = (int)GlobalSize(locked).ToUInt32() / Marshal.SizeOf<TSpan>();
                    var span = new ReadOnlySpan<TSpan>((void*) locked, size);
                    return func(locked, span);
                }
            }
            finally
            {
                GlobalUnlock(locked);
            }
        }

        /// <summary>
        /// STGMEDIUMのHGLOBALから構造体一つ分を読み込みます
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="stg"></param>
        /// <returns></returns>
        public static TResult ReadHGlobal<TResult>(this in STGMEDIUM stg)
        {
            if (stg.tymed != TYMED.TYMED_HGLOBAL) throw new ArgumentException(nameof(stg));
            var locked = GlobalLock(stg.unionmember);
            try
            {
                var size = (int) GlobalSize(locked).ToUInt32();
                unsafe
                {
                    return new ReadOnlySpan<TResult>((void*) locked, size)[0];
                }
            }
            finally
            {
                GlobalUnlock(locked);
            }
        }


        /// <summary>
        /// 拡張メタファイルのデータをコピーしてMemoryStreamを作成します
        /// </summary>
        /// <param name="hEnhFile">拡張メタファイルのハンドル</param>
        /// <returns>作成したStream</returns>
        private static Stream CreateStreamFromEnhMetaFile(IntPtr hEnhFile)
        {
            var size = GetEnhMetaFileBits(hEnhFile, 0, IntPtr.Zero);
            var bin = new byte[size];
            GetEnhMetaFileBits(hEnhFile, size, bin);
            return new MemoryStream(bin);
        }

        /// <summary>
        /// メタファイルのデータをコピーしてMemoryStreamを作成します
        /// </summary>
        /// <param name="hMetaFile">メタファイルのハンドル</param>
        /// <returns>作成したStream</returns>
        private static Stream CreateStreamFromMetaFile(IntPtr hMetaFile)
        {
            var size = GetMetaFileBitsEx(hMetaFile, 0, IntPtr.Zero);
            var bin = new byte[size];
            GetMetaFileBitsEx(hMetaFile, size, bin);
            return new MemoryStream(bin);
        }

        public static Stream GetManagedStream(in this STGMEDIUM stg)
        {
            switch (stg.tymed)
            {
                case TYMED.TYMED_MFPICT:
                    return CreateStreamFromMetaFile(stg.unionmember);
                case TYMED.TYMED_ENHMF:
                    return CreateStreamFromEnhMetaFile(stg.unionmember);
                default:
                {
                    using var s = GetUnmanagedStream(stg, true);
                    s.Seek(0, SeekOrigin.Begin);
                    var m = new MemoryStream((int) s.Length);
                    s.CopyTo(m);
                    m.Position = 0;
                    return m;
                }
            }
        }

        /// <summary>
        /// Unmanagedなメモリ領域を持つStreamを作成します。
        /// 2回目以降の呼び出し結果は未定義となります。
        /// </summary>
        /// <param name="stg">IStreamまたはHGLOBALであるSTGMEDIUM</param>
        /// <param name="autoRelease">作成したStreamが破棄された場合、stg自体を破棄するかを指定します。
        /// trueを指定した場合、呼び出し元は作成したStreamのDispose以外の手段でメモリを開放してはいけません。</param>
        /// <returns></returns>
        public static Stream GetUnmanagedStream(this STGMEDIUM stg, bool autoRelease)
        {
            IStream s;
            switch (stg.tymed)
            {
                case TYMED.TYMED_HGLOBAL: // create IStream
                    CreateStreamOnHGlobal(stg.unionmember, false, out s).ThrowIfFailed();
                    break;
                case TYMED.TYMED_ISTREAM: // cast to IStream
                    s = (IStream) Marshal.GetObjectForIUnknown(stg.unionmember);
                    break;
                default: // Error
                    throw new NotImplementedException(stg.tymed.ToString());
            }

            var cs = new ComStream(s, false, true);

            // Release StgMedium
            if (autoRelease)
                cs.Disposed += (sender, __) => // when ComStream disposed
                {
                    // release created IStream
                    if ((stg.tymed & TYMED.TYMED_HGLOBAL) != 0) Marshal.ReleaseComObject(s);
                    stg.Dispose();
                };
            return cs;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static FormatId GetFormatId(in this FORMATETC f)
            => new FormatId(f.cfFormat);
    }

    public enum NativeStringType
    {
        Unicode,
        Ansi,
        Utf8
    }
}