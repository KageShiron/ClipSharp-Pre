using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DataFormatLib;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace DataFormatLib
{
    class ComStream : Stream
    {
        private readonly IStream _stream;
        private readonly bool _autoRelease;
        private readonly bool _readOnly;
        public IStream BaseStream => _stream;

        #region Stream

        /// <summary>
        /// Initialize ComStream Object
        /// </summary>
        /// <param name="istream"></param>
        /// <param name="autoReleaseAsComObject"></param>
        public ComStream(IStream istream, bool autoReleaseAsComObject = false, bool asReadOnly = false)
        {
            _stream = istream;
            _autoRelease = autoReleaseAsComObject;
            _readOnly = asReadOnly;
        }

        public override void Flush()
        {
            this._stream.Commit(0);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long pos;
            _stream.Seek(offset, origin, out pos);
            return pos;
        }

        public override void SetLength(long value)
        {
            if (_readOnly) throw new NotSupportedException();
            _stream.SetSize(value);
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            uint read;
            if (offset != 0) throw new NotImplementedException();
            _stream.Read(buffer, count, out read);
            return (int) read;
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            if(_readOnly)throw new NotSupportedException();
            if (offset != 0) throw new NotImplementedException();
            _stream.Write(buffer, count, IntPtr.Zero);
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;

        public override bool CanWrite
        {
            get
            {
                if (_readOnly) return false;
                STATSTG statstg;
                _stream.Stat(out statstg, STATFLAG.NONAME);
                // STGM_READ = 1 or STGM_READWRITE = 2
                return (statstg.grfMode & 0x3) != 0;
            }
        }

        public override long Length
        {
            get
            {
                STATSTG statstg;
                _stream.Stat(out statstg, STATFLAG.NONAME);
                // STGM_READ = 1 or STGM_READWRITE = 2
                return statstg.cbSize;
            }
        }

        public override long Position
        {
            get { return this.Seek(0, SeekOrigin.Current); }
            set { this.Seek(0, SeekOrigin.Begin); }
        }

        #endregion Stream

        #region Dispose

        public event EventHandler Disposed;

        private bool disposed = false;

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (!disposed)
                {
                    disposed = true;
                    Disposed?.Invoke(this,new EventArgs());
                    if (_autoRelease) Marshal.ReleaseComObject(_stream);

                    if (disposing)
                    {
                        GC.SuppressFinalize(this);
                    }
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }
        #endregion Dispose
    }
}