using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

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

        public ISheet this[int index] => throw new NotImplementedException();

        public ISheet this[string sheetName] => throw new NotImplementedException();

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
