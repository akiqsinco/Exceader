using Exceader.Models;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Exceader.Test
{
    public class BookTest
    {
        public static readonly string NormalData = @"data\normal.xlsx";
        public static readonly string ExceptionData = @"data\exception.xlsx";

        [Test]
        public void OpenNormalFile()
        {
            Assert.That(() =>
            {
                using (Book.Open(NormalData)) { }
            }, Throws.Nothing);
        }

        [Test]
        public void OpenExceptionFile()
        {
            Assert.That(() => 
            {
                using (Book.Open(ExceptionData)) { }
            }, Throws.Exception.TypeOf<InvalidDataException>());
        }

        [Test]
        public void OpenByNullPath()
        {
            Assert.That(() => Book.Open((string)null), Throws.ArgumentNullException);
        }

        [Test]
        public void OpenByNullStream()
        {
            Assert.That(() => Book.Open((Stream)null), Throws.ArgumentNullException);
        }

        [Test]
        public void GetBookNames()
        {
            using (var book = Book.Open(NormalData))
            {
                var names = book.SheetNames.ToList();
                Assert.That(names, Is.EquivalentTo(new[] { "TestSheet1", "TestSheet2" }));
            }
        }

        [Test]
        [TestCase("TestSheet1")]
        [TestCase("TestSheet2")]
        public void GetExistSheet(string sheetName)
        {
            using (var book = Book.Open(NormalData))
            {
                Assert.That(() => book[sheetName], Throws.Nothing);
                Assert.That(book[sheetName].Name, Is.EqualTo(sheetName));
            }
        }

        [Test]
        public void GetNotExistSheet()
        {
            using (var book = Book.Open(NormalData))
            {
                Assert.That(() => book["NotExist"], Throws.Exception.TypeOf<KeyNotFoundException>());
            }
        }

        [Test]
        public void GetSheetByNullName()
        {
            using (var book = Book.Open(NormalData))
            {
                Assert.That(() => book[null], Throws.ArgumentNullException);
            }
        }
    }
}
