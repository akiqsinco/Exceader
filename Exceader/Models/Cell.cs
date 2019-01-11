using System;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Cell : ICell
    {
        public IRow Row => throw new NotImplementedException();

        public string Value => throw new NotImplementedException();

        public int RowIndex => throw new NotImplementedException();

        public int ColumnIndex => throw new NotImplementedException();

        public string Id => throw new NotImplementedException();

        public override string ToString() => throw new NotImplementedException();
    }
}
