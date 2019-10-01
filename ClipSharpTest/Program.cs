using ClipSharp;
using System;

namespace ClipSharpTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var x = new ComDataObject();
            foreach (var item in x.GetFormats())
            {
                Console.WriteLine(item);
            }
            Console.WriteLine(x.GetString(FormatId.Html));
        }
    }
}
