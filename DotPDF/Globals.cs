using Newtonsoft.Json.Linq;

namespace DotPDF
{
    public class Globals
    {
        public dynamic Obj { get; set; }
        public dynamic Item { get; set; }
        public JArray Array { get; set; }

        public bool IsEven(JArray array)
        {
            return array.IndexOf(Item) % 2 == 0;
        }

        public int IndexOf(JArray array, dynamic obj)
        {
            for (var i = 0; i < array.Count; i++)
            {
                var e = array[i];
                if (e == obj)
                    return i;
            }
            return -1;
        }
    }
}
