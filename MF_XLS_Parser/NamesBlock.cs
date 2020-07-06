using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MF_XLS_Parser
{
    public class NamesBlock
    {
        public string Section { get; }
        public string Group { get; }
        public string Category { get; }
        public string SubCategory { get; }
        public int StartRow { get; }
        public int EndRow { get; }

        public NamesBlock(int start, int end, string section, string group, string cat,
            string sub)
        {
            StartRow = start;
            EndRow = end;
            Section = section;
            Group = group;
            Category = cat;
            SubCategory = sub;
        }

    }
}
