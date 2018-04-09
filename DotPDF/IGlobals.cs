using Newtonsoft.Json.Linq;

namespace DotPDF
{
    public interface IGlobals
    {

        dynamic Obj { get; set; }
        dynamic Item { get; set; }
        JArray Array { get; set; }

        bool IsEven(JArray array);
        int IndexOf(JArray array, dynamic obj);

    }
}
