using Exceader.Models;
using Exceader.Common;
using NUnit.Framework;

namespace Exceader.Test
{
    public class CellTest
    {
        [Test]
        [TestCase(0)]
        [TestCase(10)]
        public void GetRowOfCell(int index)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var row = book["TestSheet1"][0];
                var cell = row[index];

                Assert.That(cell.Row, Is.EqualTo(row));
            }
        }

        [Test]
        [TestCase("A1", "Sheet1!A1")]
        [TestCase("A2", "123.0")]
        [TestCase("B1", "Sheet1!B1")]
        [TestCase("B2", "15129")]
        public void GetFilledCellValue(string id, string expected)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[id].Value, Is.EqualTo(expected));
            }
        }

        [Test]
        [TestCase("A4")]
        [TestCase("XFD65535")]
        public void GetEmptyCellValue(string id)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[id].IsEmpty);
            }
        }

        [Test]
        [TestCase("A1")]
        [TestCase("XFD1048576")]
        public void GetCellId(string id)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(sheet[id].Id, Is.EqualTo(id));
            }
        }

        [Test]
        [TestCase("A")]
        [TestCase("XFD")]
        public void GetColumnName(string columnName)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var row = book["TestSheet1"][0];

                Assert.That(row[columnName].ColumnName, Is.EqualTo(columnName));
            }
        }
    }
}
