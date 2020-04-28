using ClipSharp;
using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Vanara.PInvoke;
using static Vanara.PInvoke.Shell32;

namespace ClipSharpTest
{
    class Program
    {


        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void OleInitialize(IntPtr pvReserved);
        [STAThread()]
        static void Main(string[] args)
        {
            OleInitialize(IntPtr.Zero);
            var dx = Clipboard.GetDataObject();
            //dx.GetBitmap();
            try
            {
                dx.GetBitmap2();
            }
            catch (Exception e) { }

            dx.GetBitmap3();
            ////dx.GetStream(FormatId.FromName("Art::GVML ClipFormat")).CopyTo(File.OpenWrite(@"D:\temp\hoge.zip"));
            var d = new DataStore();
            //var z = new ZipArchive(d.GetData<Stream>("Art::GVML ClipFormat"), ZipArchiveMode.Read);
            
            //d.SetData(FormatId.CF_BITMAP, Image.FromFile(@"C:\Users\nagatsuki\Pictures\img008.jpg"));
            //d.SetData(FormatId.FromName("PNG"), File.OpenRead(@"D:\gd\pics\79574.png")); d.SetData(FormatId.CF_TEXT,"hoge");
            IDataObject x = d;

            Clipboard.OleSetClipboard(d);
            Ole32.OleFlushClipboard();

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

    }
}
