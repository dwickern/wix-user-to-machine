using System.IO;
using System.Linq;
using System.Collections.Generic;
using System;

namespace UninstallRelatedProducts
{
    public class Logger : IDisposable
    {
        readonly StreamWriter _writer;

        public Logger(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                _writer = File.CreateText(path);
            }
        }

        public void Log(object value)
        {
            Console.WriteLine(value);
            if (_writer != null)
            {
                _writer.WriteLine(value);
            }
        }

        #region Implementation of IDisposable

        public void Dispose()
        {
            if (_writer != null)
            {
                _writer.Dispose();
            }
        }

        #endregion
    }
}