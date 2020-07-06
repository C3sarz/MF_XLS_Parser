using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MF_XLS_Parser
{
    public class DataBlock
    {   
        public int StartRow { get; }
        public int EndRow { get; }
        public int Code { get; }
        public int Quantity { get; }
        public int Total { get; }

        public DataBlock(int start, int end,int code)
        {
            StartRow = start;
            EndRow = end;
        }
    }
}
