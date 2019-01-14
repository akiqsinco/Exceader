using Exceader.Models;
using NUnit.Framework;
using System;

namespace Exceader.Test
{
    public class CellPositionTest
    {
        [Test]
        [TestCase("A1", 0, 0)]
        [TestCase("XFD1048576", 1048575, 16383)]
        public void ParseValidId(string id, int expectedRow, int expectedColumn)
        {
            var position = CellPosition.Parse(id);

            Assert.That(position.Row, Is.EqualTo(expectedRow));
            Assert.That(position.Column, Is.EqualTo(expectedColumn));
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
            Assert.That(() => CellPosition.Parse(id), Throws.Exception.TypeOf<FormatException>());
        }

        [Test]
        public void ParseNullId()
        {
            Assert.That(() => CellPosition.Parse(null), Throws.ArgumentNullException);
        }

        [Test]
        [TestCase("A1", 0, 0)]
        [TestCase("XFD1048576", 1048575, 16383)]
        public void TryParseValidId(string id, int expectedRow, int expectedColumn)
        {
            var success = CellPosition.TryParse(id, out var position);

            Assert.That(success);
            Assert.That(position.Row, Is.EqualTo(expectedRow));
            Assert.That(position.Column, Is.EqualTo(expectedColumn));
        }

        [Test]
        [TestCase("12345")]
        [TestCase("ABC")]
        [TestCase("A0")]
        [TestCase("@1")]
        [TestCase("_")]
        [TestCase("")]
        public void TryParseInvalidId(string id)
        {
            var success = CellPosition.TryParse(id, out var position);

            Assert.That(!success);
            Assert.That(position, Is.EqualTo(default(CellPosition)));
        }

        [Test]
        public void TryParseNullId()
        {
            Assert.That(() => CellPosition.TryParse(null, out var pos), Throws.ArgumentNullException);
        }

        [Test]
        [TestCase(0, 0, "A1")]
        [TestCase(1048575, 16383, "XFD1048576")]
        public void ConvertToId(int row, int column, string expectedId)
        {
            var position = new CellPosition(row, column);

            Assert.That(position.ToId(), Is.EqualTo(expectedId));
        }

        [Test]
        public void ConvertDefaultValueToId()
        {
            var position = new CellPosition();

            Assert.That(position.ToId(), Is.Empty);
        }

        [Test]
        [TestCase(-1, 0)]
        [TestCase(0, -1)]
        [TestCase(-1, -1)]
        public void InstanciateWithInvalidIndex(int row, int column)
        {
            Assert.That(() => new CellPosition(row, column), Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
        }
    }
}
