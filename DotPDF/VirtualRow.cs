using System.Collections;

namespace DotPDF
{
    internal class VirtualRow
    {
        public int RowIndex { get; set; }
        public IList Items { get; set; }
    }
}
