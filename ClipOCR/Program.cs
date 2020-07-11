using System;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using System.IO;
using System.Threading.Tasks;
using ClipSharp;
using System.CommandLine;
using Windows.Globalization;
using System.CommandLine.Invocation;
using System.Net;
using Windows.Media.Capture;
using Windows.UI.Xaml;

namespace ClipOcr
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            RootCommand root = new RootCommand();
            Option lang = new Option<string>(new string[] { "--lang", "-l" }, () => "system", "");
            Option strip = new Option<bool>(new string[] { "--strip", "-s" }, () => false,"");
            root.AddOption(lang);
            root.AddOption(strip);
            root.Handler = CommandHandler.Create<string,bool>(Handler);
            await root.InvokeAsync(args);
        }

        static async Task Handler(string lang,bool strip)
        {
            var engine = lang == "system" ? OcrEngine.TryCreateFromUserProfileLanguages() : OcrEngine.TryCreateFromLanguage(new Language(lang));
            var clip = await ClipSharp.Clipboard.GetDataObject();
            var png = clip.GetStream(FormatId.FromName("PNG"));
            if (png == null) Exit("Clipboard has no valid image.");
            {
                Windows.Graphics.Imaging.BitmapDecoder decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(png.AsRandomAccessStream());
                SoftwareBitmap softwareBitmap = await decoder.GetSoftwareBitmapAsync();
                var result = await engine.RecognizeAsync(softwareBitmap);
                if (string.IsNullOrWhiteSpace(result.Text))
                {
                    Exit("No character was detected");
                }
                var d = new DataStore();
                var txt = result.Text;
                if (strip) txt = txt.Replace(" ", "");
                Console.WriteLine(txt);
                d.SetString(txt);
                await ClipSharp.Clipboard.SetClipboard(d);
            }
        }

        static void Exit(string text)
        {
            Console.Error.WriteLine(text);
            Environment.Exit(1);
        }

    }
}