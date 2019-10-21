using ClipSharp;
using System;
using System.IO;
using System.Runtime.InteropServices;
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
            var x = Clipboard.GetDataObject();
            foreach (var item in x.GetFileContents())
            {
                Console.WriteLine(item.Key.FileName);
                Console.WriteLine(new StreamReader( item.Value).ReadToEnd());

            }
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
