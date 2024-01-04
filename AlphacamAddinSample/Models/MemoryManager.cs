using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AlphacamAddinSample.Models
{
    internal class MemoryManager : IDisposable
    {
        private List<object> _comObjects;
        private bool _bDisposed;

        public MemoryManager() => this._comObjects = new List<object>();

        public T Add<T>(T item)
        {
            _comObjects.Add( item);
            return item;
        }

        public void Dispose() => this.DisposeClass();

        ~MemoryManager() => this.DisposeClass();

        private void DisposeClass()
        {
            if (this._bDisposed)
                return;
            foreach (object comObject in this._comObjects)
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                    Marshal.ReleaseComObject(comObject);
            }
            this._comObjects.Clear();
            this._comObjects =  null;
            this._bDisposed = true;
            GC.SuppressFinalize( this);
        }
    }
}