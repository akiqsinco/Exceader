using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Exceader.Models
{
    public class Row : IRow
    {
        public ICell this[int index] => throw new NotImplementedException();

        public ISheet Sheet => throw new NotImplementedException();
    }
}
