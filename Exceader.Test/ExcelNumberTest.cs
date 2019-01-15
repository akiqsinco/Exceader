using System;
using System.Collections.Generic;
using System.Text;
using Exceader.Common;
using NUnit.Framework;

namespace Exceader.Test
{
    public class ExcelNumberTest
    {
        [Test]
        [TestCase("A1", 0, 0)]
        [TestCase("XFD1048576", 1048575, 16383)]
        public void ParseValidId(string id, int expectedRow, int expectedColumn)
        {
            var (row, column) = ExcelNumber.ToIndexes(id);

            Assert.That(row, Is.EqualTo(expectedRow));
            Assert.That(column, Is.EqualTo(expectedColumn));
        }

        [Test]
        [TestCase("12345")]
        [TestCase("ABC")]
        [TestCase("A0")]
        [TestCase("@1")]
        [TestCase("_")]
        [TestCase("")]
        public void ParseInvalidId(string id)
        {
            Assert.That(() => ExcelNumber.ToIndexes(id), Throws.Exception.TypeOf<FormatException>());
        }

        [Test]
        public void ParseNullId()
        {
            Assert.That(() => ExcelNumber.ToIndexes(null), Throws.ArgumentNullException);
        }

        [Test]
        [TestCase("A", 0)]
        [TestCase("XFD", 16383)]
        public void ParseValidNumber(string alphabetical, int expected)
        {
            var numerical = ExcelNumber.ToNumerical(alphabetical);

            Assert.That(numerical, Is.EqualTo(expected));
        }

        [Test]
        [TestCase("12345")]
        [TestCase("A0")]
        [TestCase("_")]
        [TestCase("")]
        public void ParseInvalidNumber(string alphabetical)
        {
            Assert.That(() => ExcelNumber.ToNumerical(alphabetical), Throws.Exception.TypeOf<FormatException>());
        }

        [Test]
        public void ParseNullNumber()
        {
            Assert.That(() => ExcelNumber.ToNumerical(null), Throws.Exception.TypeOf<ArgumentNullException>());
        }

        [Test]
        [TestCase(0, 0, "A1")]
        [TestCase(1048575, 16383, "XFD1048576")]
        public void ConvertToId(int row, int column, string expectedId)
        {
            var id = ExcelNumber.ToId(row, column);

            Assert.That(id, Is.EqualTo(expectedId));
        }

        [Test]
        [TestCase(-1, 0)]
        [TestCase(0, -1)]
        [TestCase(-1, -1)]
        public void InstanciateWithInvalidIndex(int row, int column)
        {
            Assert.That(() => ExcelNumber.ToId(row, column), Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
        }
    }
}
