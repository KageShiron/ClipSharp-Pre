using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using IDataObject = System.Windows.Forms.IDataObject;



namespace ClipSharp
{
    public class DataObject : IComDataObject, IDataObject
    {
        private readonly IDataObject innerData;
        private readonly Dictionary<FormatId, object> store = new Dictionary<FormatId, object>();

        public DataObject(IDataObject data)
        {
            innerData = data;
        }

        public DataObject(IComDataObject data)
        {
            if (data is IDataObject dataObject)
            {
                innerData = dataObject;
            }
            else
            {
                innerData = new OleDataObject(data);
            }
        }

        public void SetData<T>(T data)
        {
            SetData(FormatId.FromDotNetName(typeof(T).FullName), data);
        }

        public void SetData(FormatId id, object data)
        {
            store[id] = data;
        }

        public T GetData<T>()
        {
            return GetData<T>(FormatId.FromDotNetName(typeof(T).FullName));
        }

        public T GetData<T>(FormatId id)
        {
            return store.TryGetValue(id, out object value) ? (T)value : throw new DirectoryNotFoundException();
        }

        public bool GetDataPresent(FormatId id)
        {
            return store.ContainsKey(id);
        }
        public bool TryGetData<T>(FormatId id, out T data)
        {
            if (store.TryGetValue(id, out object value))
            {
                if (value is T d)
                {
                    data = d;
                    return true;
                }
            }

            data = default;
            return false;
        }

        public IEnumerable<FormatId> GetFormats() => store.Keys;


        #region System.Windows.Forms.IComDataObject
        object IDataObject.GetData(string format, bool autoConvert) => GetData<object>(FormatId.FromDotNetName("format"));

        object IDataObject.GetData(string format) => GetData<object>(FormatId.FromDotNetName(format));

        object IDataObject.GetData(Type format) => GetData<object>(FormatId.FromDotNetName(format.FullName));

        void IDataObject.SetData(string format, bool autoConvert, object data) => SetData(FormatId.FromDotNetName(format), data);

        void IDataObject.SetData(string format, object data) => ((IDataObject)this).SetData(format, false, data);

        void IDataObject.SetData(Type format, object data) => ((IDataObject)this).SetData(format.FullName, false, data);

        void IDataObject.SetData(object data) => ((IDataObject)this).SetData(data.GetType().FullName, false, data);

        bool IDataObject.GetDataPresent(string format, bool autoConvert) =>
            GetDataPresent(FormatId.FromDotNetName(format));

        bool IDataObject.GetDataPresent(string format) => ((IDataObject)this).GetDataPresent(format, false);

        bool IDataObject.GetDataPresent(Type format) => ((IDataObject)this).GetDataPresent(format.FullName, false);


        string[] IDataObject.GetFormats(bool autoConvert) => this.GetFormats().Select(x => x.DotNetName).ToArray();

        string[] IDataObject.GetFormats() => this.GetFormats().Select(x => x.DotNetName).ToArray();
        #endregion

        #region IComDataObject
        int IComDataObject.DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            throw new NotImplementedException();
        }

        void IComDataObject.DUnadvise(int connection)
        {
            throw new NotImplementedException();
        }

        int IComDataObject.EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            throw new NotImplementedException();
        }

        IEnumFORMATETC IComDataObject.EnumFormatEtc(DATADIR direction)
        {
            throw new NotImplementedException();
        }

        int IComDataObject.GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            throw new NotImplementedException();
        }

        void IComDataObject.GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            throw new NotImplementedException();
        }

        void IComDataObject.GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            throw new NotImplementedException();
        }

        int IComDataObject.QueryGetData(ref FORMATETC format)
        {
            throw new NotImplementedException();
        }

        void IComDataObject.SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            throw new NotImplementedException();
        }

        #endregion
    }

    class DataStoreEntry<T>
    {
        public T Data { get; }

        public DataStoreEntry(T data)
        {
            this.Data = data;
        }
    }
}
