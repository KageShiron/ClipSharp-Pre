using ClipSharp;
using System;

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
    }
}
