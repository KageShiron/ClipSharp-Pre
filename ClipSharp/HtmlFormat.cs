using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ClipSharp
{
    public class HtmlFormat
    {
        static public HtmlFormat Parse(string val)
        {
            var html = new HtmlFormat();
            var lines = val.Split(new string[] { "\n", "\n\r" }, StringSplitOptions.RemoveEmptyEntries);
            int i;
            for (i = 0; i < lines.Length; i++)
            {
                var l = lines[i].Trim();
                if (l.StartsWith("<")) break;
                else if (l.StartsWith("Version"))
                {
                    html.Version = Regex.Match(l, @"Version\s*:\s*(.*?)$").Groups[1].Value;
                }
                else if (l.StartsWith("StartHTML"))
                {
                    html.StartHtml = int.Parse(Regex.Match(l, @"StartHTML\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("EndHTML"))
                {
                    html.EndHtml = int.Parse(Regex.Match(l, @"EndHTML\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("StartFragment"))
                {
                    html.StartFragment = int.Parse(Regex.Match(l, @"StartFragment\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("EndFragment"))
                {
                    html.EndFragment = int.Parse(Regex.Match(l, @"EndFragment\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("StartSelection"))
                {
                    html.StartSelection = int.Parse(Regex.Match(l, @"StartSelection\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("EndSelection"))
                {
                    html.EndSelection = int.Parse(Regex.Match(l, @"EndSelection\s*:\s*(.*?)$").Groups[1].Value);
                }
                else if (l.StartsWith("SourceURL"))
                {
                    html.SourceUrl = Regex.Match(l, @"SourceURL\s*:\s*(.*?)$").Groups[1].Value;
                }
            }
            html.Fragment = Regex.Match(val, @"<!--\s*StartFragment\s*-->(.*?)<!--\s*EndFragment\s*-->", RegexOptions.Singleline).Groups[1].Value;
            html.Html = string.Join("\n", lines.AsSpan().Slice(i).ToArray());
            return html;
        }

        public string Version { get; set; }
        public int StartHtml { get; set; } = -1;
        public int EndHtml { get; set; } = -1;
        public int StartFragment { get; set; } = -1;
        public int EndFragment { get; set; } = -1;
        public int StartSelection { get; set; } = -1;
        public int EndSelection { get; set; } = -1;

        public string SourceUrl { get; set; }

        public string Html { get; set; }

        public string Fragment { get; set; }
    }
}
