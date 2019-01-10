using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Exceader.Models
{
    public class Book : IBook
    {
        public static IBook Open(string path)
        {
            throw new NotImplementedException();
        }

        public static IBook Open(Stream stream)
        {
            throw new NotImplementedException();
        }

        private readonly ISheet[] _sheets;

        public ISheet this[int index] => throw new NotImplementedException();

        public ISheet this[string sheetName] => throw new NotImplementedException();

        public int Count => _sheets.Length;

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public IEnumerator<ISheet> GetEnumerator()
        {
            return ((IEnumerable<ISheet>)_sheets).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _sheets.GetEnumerator();
        }
    }
}
