using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Excel = Microsoft.Office.Interop.Excel;

namespace Finalizers
{
    public class MixedClass : IDisposable
    {
        private StreamWriter _writer;
        private Excel.Application _excel;

        private bool disposedValue;

        public void StartWriting()
        {
            _writer = new StreamWriter("output.txt");
            _excel = new Excel.Application();
        }

        [SupportedOSPlatform("windows")]
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    _writer?.Dispose();
                    Console.WriteLine("Disposing of writer");
                }

                if(_excel != null)
                {
                    _excel.Quit();
                    Marshal.ReleaseComObject(_excel);
                    Console.WriteLine("Releasing Excel");
                }

                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        [SupportedOSPlatform("windows")]
        ~MixedClass()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: false);
        }

        [SupportedOSPlatform("windows")]
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
