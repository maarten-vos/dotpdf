using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotPDF
{
    internal class VirtualRow
    {
        public int RowIndex { get; set; }
        public IList Kids { get; set; }
    }
}
