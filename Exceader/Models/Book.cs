using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace Exceader.Models
{
    public class Book : IBook
    {
        private readonly ZipArchive _archive;

        private readonly IReadOnlyList<string> _sharedStrings;

        private readonly IReadOnlyDictionary<string, string> _sheetEntryNames;

        private readonly IDictionary<string, ISheet> _sheets;

        public IEnumerable<string> SheetNames => _sheetEntryNames.Keys;

        private bool _disposed = false;

        public ISheet this[string sheetName]
        {
            get
            {
                if (!_sheetEntryNames.ContainsKey(sheetName))
                {
                    throw new ArgumentOutOfRangeException(nameof(sheetName));
                }

                if (!_sheets.ContainsKey(sheetName))
                {
                    _sheets[sheetName] = CreateSheet(sheetName);
                }

                return _sheets[sheetName];
            }
        }

        private Book(Stream stream)
        {
            var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            if (!IsValidArchive(archive))
            {
                throw new InvalidDataException();
            }

            _archive = archive;
            _sharedStrings = CreateSharedStrings(archive);
            _sheetEntryNames = CreateSheetEntryNames(archive);
            _sheets = new Dictionary<string, ISheet>();
        }

        ~Book()
        {
            Dispose(false);
        }

        public static IBook Open(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException(nameof(path));
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException();
            }

            var fs = new FileStream(path, FileMode.Open);
            return new Book(fs);
        }

        public static IBook Open(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            return new Book(stream);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                _archive?.Dispose();
            }

            _disposed = true;
        }

        private static bool IsValidArchive(ZipArchive archive)
        {
            return archive.GetEntry("xl/workbook.xml") != null;
        }

        private static IReadOnlyDictionary<string, string> CreateSheetEntryNames(ZipArchive archive)
        {
            var sheetNameToRid = new Dictionary<string, string>();
            var workbook = archive.GetEntry("xl/workbook.xml");
            using (var stream = workbook.Open())
            {
                var xml = new XmlDocument();
                xml.Load(stream);

                foreach (XmlElement sheet in xml.GetElementsByTagName("sheet"))
                {
                    sheetNameToRid[sheet.GetAttribute("name")] = sheet.GetAttribute("r:id");
                }
            }

            var ridToEntryName = new Dictionary<string, string>();
            var rels = archive.GetEntry("xl/_rels/workbook.xml.rels");
            using (var stream = rels.Open())
            {
                var xml = new XmlDocument();
                xml.Load(stream);

                foreach (XmlElement rel in xml.GetElementsByTagName("Relationship"))
                {
                    ridToEntryName[rel.GetAttribute("Id")] = "xl/" + rel.GetAttribute("Target");
                }
            }

            return sheetNameToRid.ToDictionary(kv => kv.Key, kv => ridToEntryName[kv.Value]);
        }

        private static IReadOnlyList<string> CreateSharedStrings(ZipArchive archive)
        {
            var sharedStrings = new List<string>();
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            using (var stream = entry.Open())
            {
                var xml = new XmlDocument();
                xml.Load(stream);

                foreach (XmlElement si in xml.GetElementsByTagName("si"))
                {
                    sharedStrings.Add(si.InnerText);
                }
            }

            return sharedStrings;
        }

        private ISheet CreateSheet(string sheetName)
        {
            var entryName = _sheetEntryNames[sheetName];
            var entry = _archive.GetEntry(entryName);
            using (var stream = entry.Open())
            {
                var xml = new XmlDocument();
                xml.Load(stream);

                return new Sheet(this, xml.DocumentElement, _sharedStrings);
            }
        }
    }
}
