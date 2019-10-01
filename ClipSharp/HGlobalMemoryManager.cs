using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace ClipSharp
{

    internal sealed class HGlobalMemoryManager : MemoryManager<byte>
    {
        private readonly int _length;
        private IntPtr _ptr;
        private int _retainedCount;
        private bool _disposed;

        public HGlobalMemoryManager(IntPtr hGlobal,int length)
        {
            _ptr = hGlobal;
            _length = length;
        }

        ~HGlobalMemoryManager()
        {
            Debug.WriteLine($"{nameof(HGlobalMemoryManager)} being finalized");
            Dispose(false);
        }

        public bool IsDisposed
        {
            get
            {
                lock (this)
                {
                    return _disposed && _retainedCount == 0;
                }
            }
        }


        public bool IsRetained
        {
            get
            {
                lock (this)
                {
                    return _retainedCount > 0;
                }
            }
        }

        public override unsafe Span<byte> GetSpan() => new Span<byte>((void*)_ptr, _length);

        public override unsafe MemoryHandle Pin(int elementIndex = 0)
        {
            if (elementIndex < 0 || elementIndex > _length) throw new ArgumentOutOfRangeException(nameof(elementIndex));

            lock (this)
            {
                if (_retainedCount == 0 && _disposed)
                {
                    throw new Exception();
                }
                _retainedCount++;
            }

            void* pointer = (void*)((byte*)_ptr + elementIndex);    // T = byte
            return new MemoryHandle(pointer, default, this);
        }

        public override void Unpin()
        {
            lock (this)
            {
                if (_retainedCount > 0)
                {
                    _retainedCount--;
                    if (_retainedCount == 0)
                    {
                        if (_disposed)
                        {
                            Marshal.FreeHGlobal(_ptr);
                            _ptr = IntPtr.Zero;
                        }
                    }
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            lock (this)
            {
                _disposed = true;
                if (_retainedCount == 0)
                {
                    Marshal.FreeHGlobal(_ptr);
                    _ptr = IntPtr.Zero;
                }
            }
        }
    }
}
