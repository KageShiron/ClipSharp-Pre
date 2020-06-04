using System;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using Vanara.PInvoke;

namespace ClipSharp
{
    public static class Clipboard
    {
        private static bool IsCurrentDataObject(ComDataObject obj)
        {
            return Ole32.OleIsCurrentClipboard(obj.DataObject) == 0;
        }

        public static ComDataObject GetDataObject()
        {
            if (!OleInitialize()) throw new ThreadStateException("OleInitialize was failed. (Is thread apartment STA?)");
            Ole32.OleGetClipboard(out var obj);
            return new ComDataObject(obj);
        }
        public static ComDataObject GetDataObjectSta()
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                return GetDataObject();
            ComDataObject? obj = null;
            var t = new Thread(() => { obj = GetDataObject(); });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            if (t.Join(100))
            {
                if (obj == null) throw new ApplicationException("Can't GetDataObject");
                return obj;
            }
            throw new ApplicationException("Timeout");
        }


        internal static bool OleInitialize() => Ole32.OleInitialize(IntPtr.Zero) == Vanara.PInvoke.HRESULT.S_OK;


        public static void SetClipboard(IDataObject dataobject)
        {
            if (!OleInitialize()) throw new ThreadStateException("OleInitialize was failed. (Is thread apartment STA?)");
            Ole32.OleSetClipboard(dataobject);
            Ole32.OleFlushClipboard();
        }

        public static void SetClipboardSta(IDataObject dataObject)
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                SetClipboard(dataObject);
            var t = new Thread(() => { SetClipboard(dataObject); });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            if (!t.Join(100)) throw new ApplicationException("Timeout");
        }
    }
}