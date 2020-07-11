using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Vanara;

namespace ClipSharp.Native
{
    /// <summary>
    /// <para>Defines the CF_HDROP clipboard format. The data that follows is a double null-terminated list of file names.</para>
    /// </summary>
    // https://docs.microsoft.com/en-us/windows/desktop/api/shlobj_core/ns-shlobj_core-_dropfiles typedef struct _DROPFILES { DWORD
    // pFiles; POINT pt; BOOL fNC; BOOL fWide; } DROPFILES, *LPDROPFILES;
    public struct DROPFILES
    {
        /// <summary>
        /// <para>Type: <c>DWORD</c></para>
        /// <para>The offset of the file list from the beginning of this structure, in bytes.</para>
        /// </summary>
        public uint pFiles;

        /// <summary>
        /// <para>Type: <c>POINT</c></para>
        /// <para>The drop point. The coordinates depend on <c>fNC</c>.</para>
        /// </summary>
        public System.Drawing.Point pt;

        /// <summary>
        /// <para>Type: <c>BOOL</c></para>
        /// <para>
        /// A nonclient area flag. If this member is <c>TRUE</c>, <c>pt</c> specifies the screen coordinates of a point in a window's
        /// nonclient area. If it is <c>FALSE</c>, <c>pt</c> specifies the client coordinates of a point in the client area.
        /// </para>
        /// </summary>
        public BOOL fNC;

        /// <summary>
        /// <para>Type: <c>BOOL</c></para>
        /// <para>
        /// A value that indicates whether the file contains ANSI or Unicode characters. If the value is zero, the file contains ANSI
        /// characters. Otherwise, it contains Unicode characters.
        /// </para>
        /// </summary>
        public BOOL fWide;

        public static DROPFILES Default { get; } = new DROPFILES { pFiles = (uint)Marshal.SizeOf<DROPFILES>(), fNC = false, pt = default, fWide = true };
    }
}
