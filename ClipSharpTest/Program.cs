using ClipSharp;
using System;

namespace ClipSharpTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var x = new ComDataObject();

            foreach (var item in x.GetFileDropList())
            {
                Console.WriteLine(item);

            }
        }
    }
}
