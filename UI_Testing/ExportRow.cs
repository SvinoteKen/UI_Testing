using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UI_Testing
{
    public class ExportRow
    {
        public string TaskNumber { get; set; }
        public string ScopeType { get; set; }
        public string IssueKeyLink { get; set; }
        public string Summary { get; set; }
        public string TestCount { get; set; }
        public string StepsCount { get; set; }
        public string ReviewStatus { get; set; }
        public string Priority { get; set; }
        public string IssueStatus { get; set; }
        public string TestStatus { get; set; }
        public string Tester { get; set; }
    }
}
