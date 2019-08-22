using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelMakerConsole
{
    class DataColumn
    {
        public double[] data;
        public Dataset dataSetInfo;
        public DataColumn(JToken configuration)
        {
            dataSetInfo = new Dataset(configuration);
            JArray entries = (JArray)configuration["data"];
            List<double> dataList = new List<double>();
            foreach (JToken entry in entries)
            {
                dataList.Add((double)entry);
            }
            data = dataList.ToArray();
        }

    }
}
