using ClipSharp;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using Vanara.PInvoke;
using static Vanara.PInvoke.Shell32;

namespace ClipSharpTest
{
    class Program
    {

        static async void Main(string[] args)
        {
            var c = await Clipboard.GetDataObject();
           
            c.GetBitmap(BitmapMode.Normal).Save("d:\\temp\\normal.png");
            c.GetBitmap(BitmapMode.Bitmap).Save("d:\\temp\\bitmap.png");
            c.GetBitmap(BitmapMode.Dib).Save("d:\\temp\\dib.png");

            return;
            //return;
            ////dx.GetStream(FormatId.FromName("Art::GVML ClipFormat")).CopyTo(File.OpenWrite(@"D:\temp\hoge.zip"));
            var d = new DataStore();
            //var z = new ZipArchive(d.GetData<Stream>("Art::GVML ClipFormat"), ZipArchiveMode.Read);

            d.SetData(FormatId.CF_BITMAP, Image.FromFile(@"C:\Users\nagatsuki\Pictures\img008.jpg"));
            d.SetData(FormatId.FromName("PNG"), File.OpenRead(@"D:\gd\pics\79574.png"));
            d.SetData(FormatId.CF_TEXT, "hoge");
            IDataObject x = d;

            await Clipboard.SetClipboard(d);

            Console.WriteLine();
        }

        static void Test(int size)
        {
            for (int i = 0; i < 10000; i++)
            {
                Span<byte> buff = stackalloc byte[1000];
                // Do Something
            }
        }

        static void Heap(string s)
        {
            for (int i = 2; i < s.Length; i += 7)
            {
                int.Parse(s.Substring(i, 4));
            }
        }

        static void Stack(string s)
        {
            var st = s.AsSpan();
            for (int i = 2; i < s.Length; i += 7)
            {
                int.Parse(st.Slice(i, 4));
            }
        }
Bitmap? GetTransparenctBitmapFromHbitmap(IntPtr hBitmap)
{
    Bitmap? tempBmp = null;
    BitmapData? bmpData = null;
    Bitmap? dBmp = null;
    try
    {
        // tempBmpはhBitmapに依存しないので開放可能なはず
        tempBmp = Image.FromHbitmap(hBitmap);

        // 色数が32より小さいならこの方法では非対応 (検証不足)
        if (Image.GetPixelFormatSize(tempBmp.PixelFormat) < 32) return null;

        Rectangle bmBounds = new Rectangle(0, 0, tempBmp.Width, tempBmp.Height);
        bmpData = tempBmp.LockBits(bmBounds, ImageLockMode.ReadOnly, tempBmp.PixelFormat);

        // dBmpはbmpDataそしてtempBmpに依存
        // Format32bppPArgbが重要
        dBmp = (Bitmap)new Bitmap(bmpData.Width, bmpData.Height, bmpData.Stride, PixelFormat.Format32bppPArgb, bmpData.Scan0);

        // Cloneして返す。その際にFormat32bppArgbを指定する。
        return dBmp.Clone(bmBounds, PixelFormat.Format32bppArgb);
    }
    catch (Exception e)
    {
        // Some error
    }
    finally
    {
        // 依存関係に基づいて開放
        dBmp?.Dispose();
        if (bmpData != null) tempBmp?.UnlockBits(bmpData);
        tempBmp?.Dispose();
    }
    //失敗
    return null;
}

    }

  }