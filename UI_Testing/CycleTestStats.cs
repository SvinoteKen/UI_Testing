using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UI_Testing
{
    public class CycleTestStats
    {
        public string Url { get; set; }
        public string CycleName { get; set; }
        public string Version { get; set; }
        public string Project { get; set; }
        public int Total { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Errored { get; set; }
        public int Blocked { get; set; }
    }    
    
}
