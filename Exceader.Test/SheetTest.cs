using Exceader.Models;
using Exceader.Common;
using NUnit.Framework;
using System;

namespace Exceader.Test
{
    public class SheetTest
    {
        [Test]
        public void GetBookOfSheet()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet.Book, Is.EqualTo(book));
            }
        }

        [Test]
        [TestCase(0)]
        [TestCase(10)]
        public void GetRowByPositiveIndex(int index)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[index], Is.Not.Null);
                Assert.That(sheet[index].Index, Is.EqualTo(index));
            }
        }

        [Test]
        public void GetRowByNegativeIndex()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(() => sheet[-1], Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
            }
        }

        [Test]
        [TestCase(0, 0)]
        [TestCase(10, 10)]
        public void GetCellByValidId(int row, int column)
        {
            var id = ExcelNumber.ToId(row, column);

            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[id], Is.Not.Null);
                Assert.That(sheet[id].Row.Index, Is.EqualTo(row));
                Assert.That(sheet[id].Index, Is.EqualTo(column));
            }
        }

        [Test]
        public void GetCellByInvalidId()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(() => sheet["@"], Throws.Exception.TypeOf<FormatException>());
            }
        }

        [Test]
        public void GetCellByNullId()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(() => sheet[null], Throws.ArgumentNullException);
            }
        }

        [Test]
        [TestCase(0, 0)]
        [TestCase(10, 10)]
        public void GetCellByPositiveIndex(int row, int column)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[row, column], Is.Not.Null);
                Assert.That(sheet[row, column].Row.Index, Is.EqualTo(row));
                Assert.That(sheet[row, column].Index, Is.EqualTo(column));
            }
        }

        [Test]
        [TestCase(-1, 0)]
        [TestCase(0, -1)]
        [TestCase(-1, -1)]
        public void GetCellByNegativeIndex(int row, int column)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(() => sheet[row, column], Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
            }
        }

        [Test]
        public void CheckSheetIsEmpty()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                Assert.That(!book["TestSheet1"].IsEmpty);
                Assert.That(book["TestSheet2"].IsEmpty);
            }
        }
    }
}
