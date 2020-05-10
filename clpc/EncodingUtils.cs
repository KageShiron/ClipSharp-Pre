using System;
using System.Runtime.InteropServices;
using System.Text;

namespace clpc
{
    class EncodingUtils
    {
        private static readonly IMultiLanguage2 MultiLang = (IMultiLanguage2)new MultiLanguage();
        static EncodingUtils()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static Encoding DetectFromBom(ReadOnlySpan<byte> buff)
        {
            if (buff.Length < 2)
            {
                return null;
            }
            if (buff[0] == 0xFE && buff[1] == 0xFF) return Encoding.BigEndianUnicode;
            else if (buff[0] == 0xFF && buff[1] == 0xFE)
            {
                if (buff.Length < 4 || buff[2] != 0 || buff[3] != 0) return Encoding.Unicode;
                else return Encoding.UTF32;
            }
            else if (buff.Length >= 3 && buff[0] == 0xEF && buff[1] == 0xBB && buff[2] == 0xBF)
            {
                return Encoding.UTF8;
            }
            else if (buff.Length >= 4 && buff[0] == 0 && buff[1] == 0 &&
                buff[2] == 0xFE && buff[3] == 0xFF)
            {
                return new UTF32Encoding(bigEndian: true, byteOrderMark: true);
            }

            return null;
        }

        public static Encoding DetectEncoding(ReadOnlySpan<byte> bytes)
        {
            var enc = DetectFromBom(bytes);
            if (enc != null) return enc;
            unsafe
            {
                fixed (byte* ptr = bytes)
                {
                    int size = bytes.Length;
                    int scores = 1;
                    MultiLang.DetectInputCodepage(4, 0, ptr, ref size, out var info, ref scores);
                    return scores > 0 ? Encoding.GetEncoding(unchecked((int)info.nCodePage)) : Encoding.UTF8;
                }
            }
        }

        private struct DetectEncodingInfo
        {
            public UInt32 nLangID;
            public UInt32 nCodePage;
            public Int32 nDocPercent;
            public Int32 nConfidence;
        };

        [ComImport, Guid("275c23e2-3747-11d0-9fea-00aa003f8646")]
        private class MultiLanguage
        {
        }

        [Guid("DCCFC164-2B38-11D2-B7EC-00C04F8F5D9A"),
         InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private unsafe interface IMultiLanguage2
        {
            void GetNumberOfCodePageInfo();
            void GetCodePageInfo();
            void GetFamilyCodePage();
            void EnumCodePages();
            void GetCharsetInfo();
            void IsConvertible();
            void ConvertString();
            void ConvertStringToUnicode();
            void ConvertStringFromUnicode();
            void ConvertStringReset();
            void GetRfc1766FromLcid();
            void GetLcidFromRfc1766();
            void EnumRfc1766();
            void GetRfc1766Info();
            void CreateConvertCharset();
            void ConvertStringInIStream();
            void ConvertStringToUnicodeEx();
            void ConvertStringFromUnicodeEx();
            void DetectCodepageInIStream();
            void DetectInputCodepage(
                [In] UInt32 dwFlag,
                [In] UInt32 dwPrefWinCodePage,
                [In] byte* pSrcStr,
                [In, Out] ref Int32 pcSrcSize,
                [Out] out DetectEncodingInfo lpEncoding,
                [In, Out] ref Int32 pnScores);
            void ValidateCodePage();
            void GetCodePageDescription();
            void IsCodePageInstallable();
            void SetMimeDBSource();
            void GetNumberOfScripts();
            void EnumScripts();
            void ValidateCodePageEx();
        }

    }
}

