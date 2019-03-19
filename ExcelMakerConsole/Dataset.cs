using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;


namespace ExcelMakerConsole
{
    class Dataset
    {
        public String dataName; //(Stress/Displacement)
        public String dataUnits;
        public String dataType = ""; //(True/Engineering)
        

        public Dataset(JToken description)
        {
            dataName = (string)description["name"];
            dataUnits = (string)description["unit"];
            dataType = (string)description["engineering_or_true"];
        }
    }
}
