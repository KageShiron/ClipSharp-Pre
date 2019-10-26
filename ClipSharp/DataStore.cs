using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using IDataObject = System.Windows.Forms.IDataObject;
using IComDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;



namespace ClipSharp
{
    public class DataObject : IComDataObject,IDataObject
    {
        private readonly IDataObject innerData;

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

        #region System.Windows.Forms.IComDataObject
        object IDataObject.GetData(string format, bool autoConvert)
        {
            throw new NotImplementedException();
        }

        object IDataObject.GetData(string format)
        {
            throw new NotImplementedException();
        }

        object IDataObject.GetData(Type format)
        {
            throw new NotImplementedException();
        }

        void IDataObject.SetData(string format, bool autoConvert, object data)
        {
            throw new NotImplementedException();
        }

        void IDataObject.SetData(string format, object data)
        {
            throw new NotImplementedException();
        }

        void IDataObject.SetData(Type format, object data)
        {
            throw new NotImplementedException();
        }

        void IDataObject.SetData(object data)
        {
            throw new NotImplementedException();
        }

        bool IDataObject.GetDataPresent(string format, bool autoConvert)
        {
            throw new NotImplementedException();
        }

        bool IDataObject.GetDataPresent(string format)
        {
            throw new NotImplementedException();
        }

        bool IDataObject.GetDataPresent(Type format)
        {
            throw new NotImplementedException();
        }

        string[] IDataObject.GetFormats(bool autoConvert)
        {
            throw new NotImplementedException();
        }

        string[] IDataObject.GetFormats()
        {
            throw new NotImplementedException();
        }
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
}
