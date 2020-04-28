using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ClipSharp
{
    public static class Clipboard
    {
        [DllImport("ole32.dll")]
        private static extern int OleIsCurrentClipboard(IDataObject pDataObject);

        [DllImport("ole32.dll")]
        public static extern int OleGetClipboard(out IDataObject pDataObject);


        [DllImport("ole32.dll")]
        public static extern int OleSetClipboard(IDataObject pDataObject);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,EntryPoint = "GetTempPathW")]
        static extern int GetTempPath1(uint nBufferLength,StringBuilder sb);

        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,EntryPoint = "GetTempPathW")]
        static extern int GetTempPath2(uint nBufferLength,ref char sb);

        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,EntryPoint = "GetTempPathW")]
        static extern unsafe int GetTempPath3(uint nBufferLength,char * sb);

        private static bool IsCurrentDataObject(ComDataObject obj)
        {
            return OleIsCurrentClipboard(obj.DataObject) == 0;
        }

        public static ComDataObject GetDataObject()
        {
            OleGetClipboard(out var obj);
            return new ComDataObject(obj);
        }
    }
}