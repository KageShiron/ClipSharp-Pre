using System;
using System.Collections.Generic;
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


        private static bool IsCurrentDataObject( ComDataObject obj )
        {
            return OleIsCurrentClipboard(obj.DataObject) == 0;
        }

        public static ComDataObject GetDataObject()
        {
            OleGetClipboard(out IDataObject obj);
            return new ComDataObject(obj);
        }
    }
}
