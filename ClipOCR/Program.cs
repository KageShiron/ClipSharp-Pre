using System;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using System.IO;
using System.Threading.Tasks;
using ClipSharp;

namespace ClipOcr
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            foreach (var item in OcrEngine.AvailableRecognizerLanguages)
            {
                Console.WriteLine(item);
            }

           var clip = ClipSharp.Clipboard.GetDataObjectSta();
            var png = clip.GetStream(FormatId.FromName("PNG"));
            if (png == null) return;
            var result = test(png);
            result.Wait();
            var d = new DataStore();
            Console.WriteLine(result.Result);
            d.SetString(result.Result);
            ClipSharp.Clipboard.SetClipboard(d);
        }
        static async Task<string> test(Stream png)
        {
            var en = new Windows.Globalization.Language("en");
            var engine = OcrEngine.TryCreateFromLanguage(en);
            using (var stream = new Windows.Storage.Streams.InMemoryRandomAccessStream())
            {
                Windows.Graphics.Imaging.BitmapDecoder decoder = await Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(png.AsRandomAccessStream());
                SoftwareBitmap softwareBitmap = await decoder.GetSoftwareBitmapAsync();
                var result = await engine.RecognizeAsync(softwareBitmap);
                return (result.Text);
            }
        }
    }
}