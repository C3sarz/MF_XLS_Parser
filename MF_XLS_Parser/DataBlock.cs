using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MF_XLS_Parser
{
    public class DataBlock
    {
        public string Name;
        public long Code;
        public double Quantity;
        public double Total;
        public int DataRow;

        public DataBlock(string Name, long Code, double Quantity, double Total, int DataRow)
        {
            this.Name = Name;
            this.Code = Code;
            this.Quantity = Quantity;
            this.Total = Total;
            this.DataRow = DataRow;
        }

    }
}
