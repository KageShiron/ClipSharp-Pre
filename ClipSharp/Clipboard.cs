using System;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Schedulers;
using Vanara.PInvoke;

namespace ClipSharp
{
    public static class Clipboard
    {

        private static StaTaskScheduler _Sta = new StaTaskScheduler(1);
        private static bool IsCurrentDataObject(ComDataObject obj)
        {
            return Ole32.OleIsCurrentClipboard(obj.DataObject) == 0;
        }

        public static async Task<ComDataObject> GetDataObject()
        {
            return await Task.Factory.StartNew(() =>
            {
                if (!OleInitialize()) throw new ThreadStateException("OleInitialize was failed. (Is thread apartment STA?)");
                Ole32.OleGetClipboard(out var obj);
                return new ComDataObject(obj);
            }, CancellationToken.None, TaskCreationOptions.None, _Sta);
        }

        internal static bool OleInitialize()
        {
            var r = Ole32.OleInitialize(IntPtr.Zero);
            return r == Vanara.PInvoke.HRESULT.S_OK || r == Vanara.PInvoke.HRESULT.S_FALSE;
        }


        public static async Task SetClipboard(IDataObject dataobject)
        {
            await Task.Factory.StartNew(() =>
            {
                if (!OleInitialize()) throw new ThreadStateException("OleInitialize was failed. (Is thread apartment STA?)");
                Ole32.OleSetClipboard(dataobject);
                Ole32.OleFlushClipboard();
            }, CancellationToken.None, TaskCreationOptions.None, _Sta);
        }
    }
}