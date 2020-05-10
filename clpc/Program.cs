using System;
using System.Buffers;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Linq;
using System.Threading.Tasks;
using ClipSharp;
using Vanara;

namespace clpc
{

    public enum DataType 
    {
        Auto,
        Text,
        Image,
        Png,
        Jpeg,
        Gif,
    }

    class Program
    {
        [STAThread()]
        public static void Main(string[] args)
        {
            var root = new RootCommand();
            var arg = new Argument<string>("file", () => null);
            root.AddArgument(arg);
            root.AddOption(new Option<DataType>("-t", () => DataType.Auto));
            root.Handler = CommandHandler.Create<DataType,string>(Entry);
            var res = root.Parse(args);
            var t = res.ValueForOption<DataType>("t");
            var tokens = res.FindResultFor(arg).Tokens;
            string file = tokens.FirstOrDefault()?.Value;
            Entry(t, file);
        }



        private static void SetText(string file)
        {
            var d = new DataStore();
            var stream = File.OpenRead(file);
            Span<byte> bytes = stackalloc byte[4096];
            stream.Read(bytes);
            stream.Seek(0,SeekOrigin.Begin);
            var enc = EncodingUtils.DetectEncoding(bytes);

            using var sr = new StreamReader(stream);
            d.SetString(sr.ReadToEnd());
            Clipboard.SetClipboard(d);

        }

        private static void SetImage(string file, params string[] data)
        {
            var f = File.OpenRead(file);
            var d = new DataStore();
            foreach (var dt in data)
            {
                d.SetData(dt,f);
            }

            d.SetData(FormatId.CF_BITMAP,Image.FromStream(f));
            Clipboard.SetClipboard(d);

        }


        public static void Entry( DataType t = DataType.Auto,  string file = null)
        {
            if (file == null)
            {
                var d = new DataStore();
                d.SetString(Console.In.ReadToEnd());
                Clipboard.SetClipboard(d);
            }
            else
            {
                switch (t)
                {
                    case DataType.Auto:
                        break;
                    case DataType.Text:
                        SetText(file);
                        break;
                    case DataType.Image:
                        SetImage(file);
                        break;
                    case DataType.Png:
                        SetImage(file, "PNG", "image/png");
                        break;
                    case DataType.Jpeg:
                        SetImage(file, "JFIF", "image/jpeg");
                        break;
                    case DataType.Gif:
                        SetImage(file,"GIF","image/gif");
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(t), t, null);
                }
            }


        }

    }
}
