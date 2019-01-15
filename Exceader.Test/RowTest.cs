using Exceader.Models;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Test
{
    public class RowTest
    {
        [Test]
        [TestCase(0)]
        [TestCase(10)]
        public void GetSheetOfRow(int index)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];
                var row = sheet[index];

                Assert.That(row.Sheet, Is.EqualTo(sheet));
            }
        }

        [Test]
        [TestCase(0)]
        [TestCase(10)]
        public void GetCellByPositiveIndex(int index)
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var row = book["TestSheet1"][0];

                Assert.That(row[index], Is.Not.Null);
                Assert.That(row[index].Index, Is.EqualTo(index));
            }
        }

        [Test]
        public void GetCellByNegativeIndex()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var row = book["TestSheet1"][0];

                Assert.That(() => row[-1], Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
            }
        }

        [Test]
        public void CheckRowIsEmpty()
        {
            using (var book = Book.Open(BookTest.NormalData))
            {
                var sheet = book["TestSheet1"];

                Assert.That(!sheet[0].IsEmpty);
                Assert.That(sheet[10].IsEmpty);
            }
        }
    }
}
