using ClipSharp;
using System;
using static Vanara.PInvoke.Shell32;

namespace ClipSharpTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var x = new ComDataObject();
            foreach (var item in x.GetPidl())
            {
                Console.WriteLine(item.ToString(Vanara.PInvoke.Shell32.SIGDN.SIGDN_NORMALDISPLAY));
                Console.WriteLine(item.ToString());

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
