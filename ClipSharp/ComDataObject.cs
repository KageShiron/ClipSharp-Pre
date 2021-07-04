
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Vanara.PInvoke;

namespace ClipSharp
{
    public class ComEventArgs : EventArgs
    {
        public Exception Exception { get; }
        public string MethodName { get; }

        public FormatId? FormatId { get; }

        public ComEventArgs(Exception e, FormatId? id = null, [CallerMemberName] string method = "")
        {
            Exception = e;
            FormatId = id;
            MethodName = method;
        }
    }

    public class ComDataObject : System.Windows.Forms.IDataObject
    {
        public event EventHandler<ComEventArgs> ExceptionRaised;

        //public static ILogger<ComDataObject> Logger = //LoggerFactory.Create(builder => { builder.AddConsole().AddDebug(); }).CreateLogger<ComDataObject>();
        public ComDataObject(IDataObject data) => DataObject = data;

        public IDataObject DataObject { get; }


        public HtmlFormat? GetHtml()
        {
            try
            {
                if (!GetDataPresent(FormatId.Html) && GetString(FormatId.Html) is string s) return HtmlFormat.Parse(s);
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            return null;
        }

        public string[]? GetFileDropList()
        {
            var f = FormatId.CF_HDROP.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(FormatId.CF_HDROP)) return null;
                DataObject.GetData(ref f, out s);
                return s.GetFiles();
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }


        #region GetBitmap

        public Bitmap? GetTransparentBitmap()
        {
            var f = GetFormatIds();
            if (f.Contains(FormatId.FromName("PNG")) && this.GetStream(FormatId.FromName("PNG")) is Stream st1) return new Bitmap(st1);
            if (f.Contains(FormatId.FromName("image/png")) && this.GetStream(FormatId.FromName("image/png")) is Stream st2) return new Bitmap(st2);
            if (f.Contains(FormatId.CF_HDROP) && this.GetFileDropList() is string[] strs) return new Bitmap(strs[0]);
            if (f.Contains(FormatId.CF_BITMAP) && this.GetBitmap(BitmapMode.Bitmap) is Bitmap bmp1) return bmp1;
            if (f.Contains(FormatId.CF_DIB) && this.GetBitmap(BitmapMode.Dib) is Bitmap bmp2) return bmp2;
            if (this.GetBitmap(BitmapMode.Bitmap) is Bitmap bmp3) return bmp3;
            if (this.GetBitmap(BitmapMode.Dib) is Bitmap bmp4) return bmp4;
            return null;
        }

        public Bitmap? GetBitmap(BitmapMode mode)
        {
            return mode switch
            {
                BitmapMode.Normal => GetBitmapNormal(),
                BitmapMode.Bitmap => GetBitmapBitmap(),
                BitmapMode.Dib => GetBitmapDib(),
                _ => null
            };
        }

        private Bitmap? GetBitmapNormal()
        {
            var f = FormatId.CF_BITMAP.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!this.GetDataPresent(ref f)) return null;
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                return Image.FromHbitmap(s.unionmember);
                //var ret = (Bitmap)bmp.Clone();
                //bmp.Dispose();
                //return ret;
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }



        //private Bitmap GetCF_BITMAP()
        //{
        //    IntPtr hBitmap = GetClipboardData(CF_BITMAP);
        //    if (hBitmap != IntPtr.Zero)
        //    {
        //        Bitmap bmp = Bitmap.FromHbitmap(hBitmap);
        //        Bitmap result = (Bitmap)bmp.Clone();
        //        bmp.Dispose();
        //        return bmp;
        //    }
        //}
        public T? InvokeDengerousHBitmap<T>(Func<HBITMAP, T> func)
        {
            var f = FormatId.CF_BITMAP.GetFormatEtc();
            f.tymed = TYMED.TYMED_GDI;
            STGMEDIUM s = default;
            try
            {
                if (!this.GetDataPresent(ref f)) return default(T);
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");
                return func(s.unionmember);
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return default(T);
        }

        private Bitmap? GetBitmapBitmap()
        {
            var f = FormatId.CF_BITMAP.GetFormatEtc();
            STGMEDIUM s = default;
            Bitmap? tempBmp = null;
            BitmapData? bmpData = null;
            Bitmap? dBmp = null;
            try
            {
                if (!this.GetDataPresent(ref f)) return null;
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_GDI) throw new ApplicationException("Invalid Tymed");

                // s.unionmemberからディープコピーするので、開放可能・・・なはず。
                tempBmp = Image.FromHbitmap(s.unionmember);

                if (Image.GetPixelFormatSize(tempBmp.PixelFormat) < 32) return tempBmp;

                Rectangle bmBounds = new Rectangle(0, 0, tempBmp.Width, tempBmp.Height);
                bmpData = tempBmp.LockBits(bmBounds, ImageLockMode.ReadOnly, tempBmp.PixelFormat);

                // tempBmpに依存
                dBmp = (Bitmap)new Bitmap(
                    bmpData.Width,
                    bmpData.Height,
                    bmpData.Stride,
                    PixelFormat.Format32bppPArgb,
                    bmpData.Scan0);

                return dBmp.Clone(bmBounds, PixelFormat.Format32bppArgb);
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                dBmp?.Dispose();
                if (bmpData != null) tempBmp?.UnlockBits(bmpData);
                tempBmp?.Dispose();
                s.Dispose();
            }
            return null;
        }


        private Bitmap? GetBitmapDib()
        {
            var f = FormatId.CF_DIB.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!this.GetDataPresent(ref f)) return null;
                DataObject.GetData(ref f, out s);
                if (s.tymed != TYMED.TYMED_HGLOBAL) throw new ApplicationException("Invalid Tymed");

                var result = s.InvokeHGlobal<Gdi32.BITMAPINFOHEADER, Bitmap?>((ptr, sp) =>
                {
                    unsafe
                    {
                        var bmp = sp[0];
                        if ((bmp.biCompression != Gdi32.BitmapCompressionMode.BI_BITFIELDS && bmp.biCompression != Gdi32.BitmapCompressionMode.BI_RGB) || bmp.biBitCount != 32) return null;
                        var isTopDown = bmp.biHeight < 0;
                        var height = Math.Abs(bmp.biHeight);
                        var lineSize = bmp.biBitCount / 8 * bmp.biWidth;

                        var old = new Bitmap(bmp.biWidth, bmp.biHeight, isTopDown ? lineSize : -lineSize,
                            PixelFormat.Format32bppPArgb, (IntPtr)(ptr.ToInt64() + bmp.biSize + bmp.biSizeImage
                                      + bmp.biClrUsed * Marshal.SizeOf(typeof(Gdi32.RGBQUAD))
                                      + (isTopDown ? 0 : -(int)(bmp.biSizeImage / bmp.biHeight))));
                        var dstBmp = old.Clone(new Rectangle(0, 0, old.Width, old.Height), PixelFormat.Format32bppArgb);
                        old.Dispose();
                        return dstBmp;
                    }
                });
                return result;
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }


        #endregion

        public (FORMATETC, STGMEDIUM)? GetDisposedStgMedium(FormatId id)
        {
            var f = id.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(id)) return null;
                DataObject.GetData(ref f, out s);
                return (f, s);
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e, id));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }

        public TResult? ReadHGlobal<TResult>(FormatId id) where TResult : unmanaged
        {
            var f = id.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(id)) return null;
                DataObject.GetData(ref f, out s);
                return s.ReadHGlobal<TResult>();
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e, id));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }

        public DragDropEffects? GetDragDropEffects() => ReadHGlobal<DragDropEffects>(FormatId.CFSTR_PREFERREDDROPEFFECT);

        public CultureInfo? GetCultureInfo()
        {
            var locale = ReadHGlobal<int>(FormatId.CF_LOCALE);
            if (locale == null) return null;
            return new CultureInfo(locale.Value);
        }

        public Shell32.PIDL[]? GetShellIdList(out Shell32.PIDL? parent)
        {
            var f = FormatId.CFSTR_SHELLIDLIST.GetFormatEtc();
            STGMEDIUM s = default;
            parent = null;
            try
            {
                if (!GetDataPresent(FormatId.CFSTR_SHELLIDLIST)) return null;
                DataObject.GetData(ref f, out s);

                var (list, x) = s.InvokeHGlobal<uint, (Shell32.PIDL[], Shell32.PIDL)>((p, x) =>
                  {
                      var l = new Shell32.PIDL[x[0]];
                      Shell32.PIDL? parent = null;
                      unsafe
                      {
                          byte* ptr = (byte*)p;
                          parent = new Shell32.PIDL((IntPtr)(ptr + x[1]), true, true);
                          for (var i = 1; i < x[0] + 1; i++)
                          {
                              // DO NOT release memory of child
                              var child = new Shell32.PIDL((IntPtr)(ptr + x[i + 1]), false, false);
                              var pa = new Shell32.PIDL(parent);
                              pa.Append(child);
                              l[i - 1] = pa;
                          }
                      }

                      return (l, parent);
                  });
                parent = x;
                return list;
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));

            }
            finally
            {
                s.Dispose();
            }
            return null;
        }


        public Stream? GetFileContent(int index)
        {
            try
            {
                if (!GetDataPresent(FormatId.CFSTR_FILECONTENTS))
                {
                    return GetStream(FormatId.CFSTR_FILECONTENTS, index);
                }
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            return null;
        }

        public Dictionary<FileDescriptor, Stream>? GetFileContents()
        {
            var fd = GetFileDescriptors();
            if (fd == null) return null;
            var s = new Dictionary<FileDescriptor, Stream>(fd.Length);
            for (var i = 0; i < fd.Length; i++)
            {
                var st = GetStream(FormatId.CFSTR_FILECONTENTS, i);
                if (st == null) return null;
                s.Add(fd[i], st);
            }
            return s;
        }

        public Metafile? GetMetafile()
        {
            var f = FormatId.CF_METAFILEPICT.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(FormatId.CF_METAFILEPICT)) return null;
                DataObject.GetData(ref f, out var stg);
                if (stg.tymed != TYMED.TYMED_MFPICT) throw new ApplicationException();
                return stg.InvokeHGlobal<Vanara.PInvoke.Gdi32.METAFILEPICT, Metafile>((x, y) =>
                {
                    using var meta = new Metafile(y[0].hMF.DangerousGetHandle(), new WmfPlaceableFileHeader(), false);
                    return (Metafile)meta.Clone();
                });
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }



        public Metafile? GetEnhancedMetafile()
        {
            if (!GetDataPresent(FormatId.CF_ENHMETAFILE)) return null;
            var f = FormatId.CF_ENHMETAFILE.GetFormatEtc();
            DataObject.GetData(ref f, out var stg);
            try
            {
                if (stg.tymed != TYMED.TYMED_ENHMF) throw new ApplicationException();
                return new Metafile(stg.GetManagedStream());
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                stg.Dispose();
            }
            return null;
        }


        #region GetString

        public string? GetString(FormatId id, NativeStringType type)
        {
            var f = DataObjectUtils.GetFormatEtc(id);
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(id)) return null;
                DataObject.GetData(ref f, out s);
                return s.GetString(type);
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e, id));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }

        public string? GetString() => GetString(FormatId.CF_UNICODETEXT);

        public string? GetString(FormatId id)
        {
            if (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                id == FormatId.CommaSeparatedValue)
                return GetString(id, NativeStringType.Ansi);
            if (id == FormatId.Html || id == FormatId.Xaml)
                return GetString(id, NativeStringType.Utf8);
            if (id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust)
                return GetString(id, NativeStringType.Unicode);
            throw new ArgumentException("GetString can't guess Encoding of " + id.DotNetName, nameof(id));
        }

        #endregion GetSring


        public Stream GetUnsafeUnmanagedStream(FormatId id, int lindex = -1)
        {
            var f = id.GetFormatEtc();
            f.lindex = lindex;
            DataObject.GetData(ref f, out var s);
            return s.GetUnmanagedStream(true);
        }
        public Stream? GetStream(FormatId id, int lindex = -1)
        {
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(id)) return null;
                var f = id.GetFormatEtc();
                f.lindex = lindex;
                DataObject.GetData(ref f, out s);
                return s.GetManagedStream();
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e, id));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }

        public FileDescriptor[]? GetFileDescriptors()
        {
            var f = FormatId.CFSTR_FILEDESCRIPTORW.GetFormatEtc();
            STGMEDIUM s = default;
            try
            {
                if (!GetDataPresent(FormatId.CFSTR_FILEDESCRIPTORW)) return null;
                DataObject.GetData(ref f, out s);
                return s.InvokeHGlobal<byte, FileDescriptor[]>((_, f) => FileDescriptor.FromFileGroupDescriptor(f));
            }
            catch (Exception e)
            {
                ExceptionRaised?.Invoke(this, new ComEventArgs(e));
            }
            finally
            {
                s.Dispose();
            }
            return null;
        }


        #region GetFormats
        public virtual IEnumerable<DataObjectFormat> GetFormats(bool allFormat = false)
        {
            IEnumFORMATETC enumFormatEtc = null!;
            try
            {
                enumFormatEtc = DataObject.EnumFormatEtc(DATADIR.DATADIR_GET);
                if (enumFormatEtc == null) return Array.Empty<DataObjectFormat>();
                enumFormatEtc.Reset();
                var fs = new List<DataObjectFormat>();
                if (allFormat)
                {
                    for (int i = 0; i <= 0xFFFF; i++)
                    {
                        var f = new FormatId(i);
                        var etc = f.GetFormatEtc();
                        if (new HRESULT((uint)DataObject.QueryGetData(ref etc)).Succeeded)
                            fs.Add(new DataObjectFormat(etc));
                    }
                }
                else
                {
                    var fe = new FORMATETC[1];
                    while (enumFormatEtc.Next(1, fe, null) == 0) fs.Add(new DataObjectFormat(fe[0]));
                }
                return fs;
            }
            finally
            {
                if (enumFormatEtc != null)
                    Marshal.ReleaseComObject(enumFormatEtc);
            }
        }
        string[] System.Windows.Forms.IDataObject.GetFormats()
        {
            return ((System.Windows.Forms.IDataObject)this).GetFormats(true);
        }

        string[] System.Windows.Forms.IDataObject.GetFormats(bool autoConvert)
        {
            return GetFormats().Select(x => x.ToString()).ToArray();
        }

        public IEnumerable<FormatId> GetFormatIds(bool allFormats = false)
        {
            return GetFormats(allFormats).Select(x => x.FormatId);
        }

        #endregion

        #region GetDataPresent

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(int format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        public bool GetDataPresent(string format, bool autoConvert)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetDataPresent(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(string format)
        {
            return GetDataPresent(format, true);
        }

        public bool GetDataPresent(Type format)
        {
            return GetDataPresent(format.FullName);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public bool GetDataPresent(FormatId format)
        {
            return GetDataPresent(format.Id);
        }

        public bool GetDataPresent(ref FORMATETC format)
        {
            return DataObject.QueryGetData(ref format) == 0;
        }

        #endregion GetDataPresent

        #region GetData

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object? GetData(int format) => GetData(new FormatId(format));

#pragma warning disable CS8613 // 戻り値の型における参照型の Null 許容性が、暗黙的に実装されるメンバーと一致しません。
        public object? GetData(string format, bool autoConvert) => GetData(FormatId.FromName(format));

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object? GetData(string format) => GetData(format, true);

        public object? GetData(Type format) => GetData(format.FullName);
#pragma warning restore CS8613 // 戻り値の型における参照型の Null 許容性が、暗黙的に実装されるメンバーと一致しません。

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public object? GetData(FormatId id)
        {
            if (id == FormatId.CF_TEXT || id == FormatId.Rtf || id == FormatId.CF_OEMTEXT ||
                id == FormatId.CommaSeparatedValue || id == FormatId.Html || id == FormatId.Xaml ||
                id == FormatId.CF_UNICODETEXT || id == FormatId.ApplicationTrust)
                return GetString(id);
            if (id == FormatId.CFSTR_PREFERREDDROPEFFECT) return GetDragDropEffects();
            if (id == FormatId.CF_HDROP) return GetFileDropList();
            if (id == FormatId.CFSTR_FILEDESCRIPTORW) return GetFileDescriptors();
            if (id == FormatId.CFSTR_FILENAMEW)
                return new[] { GetString(FormatId.CFSTR_FILENAMEW, NativeStringType.Unicode) };
            if (id == FormatId.CFSTR_FILENAMEA)
                return new[] { GetString(FormatId.CFSTR_FILENAMEA, NativeStringType.Ansi) };
            if (id == FormatId.CF_BITMAP) return GetBitmap(BitmapMode.Normal);
            if (id == FormatId.CF_ENHMETAFILE) return GetEnhancedMetafile();
            if (id == FormatId.CF_METAFILEPICT) return GetMetafile();
            if (id == FormatId.CFSTR_SHELLIDLIST) return GetShellIdList(out var _);
            return GetStream(id);
        }

        #endregion

        #region SetData

        public void SetData(string format, bool autoConvert, object data) => throw new NotImplementedException();

        public void SetData(string format, object data) => SetData(format, true, data);

        public void SetData(Type format, object data) => SetData(format.FullName, data);

        public void SetData(object data) => throw new NotImplementedException();
        #endregion

        #region GetCanonicalFormatEtc

        public virtual FORMATETC GetCanonicalFormatEtc(ref FORMATETC format)
        {
            DataObject.GetCanonicalFormatEtc(ref format, out var c);
            return c;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public FORMATETC GetCanonicalFormatEtc(int format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetCanonicalFormatEtc(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public FORMATETC GetCanonicalFormatEtc(string format)
        {
            var f = DataObjectUtils.GetFormatEtc(format);
            return GetCanonicalFormatEtc(ref f);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public virtual FORMATETC GetCanonicalFormatEtc(FormatId format)
        {
            return GetCanonicalFormatEtc(format.Id);
        }

        #endregion GetCanonicalFormatEtc
    }

    public enum BitmapMode
    {
        Normal,
        Bitmap,
        Dib
    }
}