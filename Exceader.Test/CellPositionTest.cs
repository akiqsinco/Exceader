using Exceader.Models;
using NUnit.Framework;
using System;

namespace Exceader.Test
{
    public class CellPositionTest
    {
        [Test]
        public void ParseMinimumIdToPosition()
        {
            var success = CellPosition.TryParse("A1", out var position);

            Assert.That(success);
            Assert.That(position.Row, Is.EqualTo(0));
            Assert.That(position.Column, Is.EqualTo(0));
        }

        [Test]
        public void ParseCommonIdToPosition()
        {
            var success = CellPosition.TryParse("XFD1048576", out var position);

            Assert.That(success);
            Assert.That(position.Row, Is.EqualTo(1048575));
            Assert.That(position.Column, Is.EqualTo(16383));
        }

        [Test]
        public void ParseRowOnlyIdToPosition()
        {
            var success = CellPosition.TryParse("12345", out var position);

            Assert.That(!success);
            Assert.That(position, Is.EqualTo(default(CellPosition)));
        }

        [Test]
        public void ParseColumnOnlyIdToPosition()
        {
            var success = CellPosition.TryParse("ABC", out var position);

            Assert.That(!success);
            Assert.That(position, Is.EqualTo(default(CellPosition)));
        }

        [Test]
        public void ParseOutOfRangeRowToPosition()
        {
            var success = CellPosition.TryParse("A0", out var position);

            Assert.That(!success);
            Assert.That(position, Is.EqualTo(default(CellPosition)));
        }

        [Test]
        public void ParseOutOfRangeColumnToPosition()
        {
            var success = CellPosition.TryParse("_0", out var position);

            Assert.That(!success);
            Assert.That(position, Is.EqualTo(default(CellPosition)));
        }

        [Test]
        public void ParseEmptyStringToPosition()
        {
            Assert.That(() => CellPosition.Parse(string.Empty), Throws.Exception.TypeOf<FormatException>());
        }

        [Test]
        public void ParseNullToPosition()
        {
            Assert.That(() => CellPosition.Parse(null), Throws.ArgumentNullException);
        }

        [Test]
        public void ConvertMinimumPositionToId()
        {
            var position = new CellPosition(0, 0);

            Assert.That(position.ToId(), Is.EqualTo("A1"));
        }

        [Test]
        public void ConvertCommonPositionToId()
        {
            var position = new CellPosition(1048575, 16383);

            Assert.That(position.ToId(), Is.EqualTo("XFD1048576"));
        }

        [Test]
        public void ConvertDefaultValueToId()
        {
            var position = new CellPosition();

            Assert.That(position.ToId(), Is.Empty);
        }

        [Test]
        public void InstanciateWithOutOrRangeRow()
        {
            Assert.That(() => new CellPosition(-1, 0), Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
        }

        [Test]
        public void InstanciateWithOutOrRangeColumn()
        {
            Assert.That(() => new CellPosition(0, -1), Throws.Exception.TypeOf<ArgumentOutOfRangeException>());
        }
    }
}
