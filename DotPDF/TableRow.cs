using Newtonsoft.Json.Linq;
using System.Collections;

namespace DotPDF
{
    internal class TableRow
    {
        public int RowIndex { get; set; }
        public IList Items { get; set; }
    }
}
