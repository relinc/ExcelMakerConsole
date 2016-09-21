using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMakerConsole
{
    class DataColumn
    {
        public double[] data;
        public Dataset dataSetInfo;
        public DataColumn(int numDataPoints)
        {
            data = new double[numDataPoints];
        }

    }
}
