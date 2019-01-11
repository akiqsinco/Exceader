using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Sheet : ISheet
    {
        public IRow this[int index] => throw new NotImplementedException();

        public ICell this[string id] => throw new NotImplementedException();

        public ICell this[int row, int column] => throw new NotImplementedException();

        public IBook Book => throw new NotImplementedException();
    }
}
