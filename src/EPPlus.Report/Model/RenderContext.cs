using System.Collections;
using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class RenderContext
    {
        public object Current { get; set; }
        public Dictionary<string, object> Variables { get; set; }
        public IEnumerable CurrentCollection { get; set; }
        public bool IsNamedRangeLoop { get; set; }
        public int CurrentIndex { get; set; }
    }
}
