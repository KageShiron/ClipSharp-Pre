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
using static Vanara.PInvoke.Shell32;
using System.Collections.Generic;

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
        Pidl,
        File
    }

    class Program
    {
        [STAThread()]
        public static void Main(string[] args)
        {
            var root = new RootCommand();
            var arg = new Argument<string[]>("files", () => null);
            root.AddArgument(arg);
            root.AddOption(new Option<DataType>("-t", () => DataType.Auto));
            root.Handler = CommandHandler.Create<DataType,string[]>(Entry);
            var res = root.Parse(args);
            var t = res.ValueForOption<DataType>("t");
            var tokens = res.FindResultFor(arg).Tokens;
            IEnumerable<string> files = tokens.Select(x => x.Value);
            Entry(t, files);
        }



        private static void SetText(string file)
        {
            var d = new DataStore();
            var stream = File.OpenRead(file);
            Span<byte> bytes = stackalloc byte[4096];
            stream.Read(bytes);

            var enc = EncodingUtils.DetectEncoding(bytes);

            stream.Seek(0,SeekOrigin.Begin);
            using var sr = new StreamReader(stream);
            var str = sr.ReadToEnd();
            d.SetString(str);

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


        public static void Entry( DataType t = DataType.Auto,  IEnumerable<string> fileList = null)
        {
            var files = fileList.ToArray();
            if (files == null)
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
                    case DataType.Text:
                        SetText(files[0]);
                        break;
                    case DataType.Image:
                        SetImage(files[0]);
                        break;
                    case DataType.Png:
                        SetImage(files[0], "PNG", "image/png");
                        break;
                    case DataType.Jpeg:
                        SetImage(files[0], "JFIF", "image/jpeg");
                        break;
                    case DataType.Gif:
                        SetImage(files[0],"GIF","image/gif");
                        break;
                    case DataType.File:
                        {
                            var d = new DataStore();
                            d.SetFileDropList(files);
                            Clipboard.SetClipboard(d);
                            break;
                        }
                    case DataType.Pidl:
                        {
                            var d = new DataStore();
                            d.SetPidl(files);
                            Clipboard.SetClipboard(d);
                            break;
                        }
                    default:
                        throw new ArgumentOutOfRangeException(nameof(t), t, null);
                }
            }


        }

    }
}
