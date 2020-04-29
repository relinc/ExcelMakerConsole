using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace ExcelMakerConsole
{
    class ReferenceSample
    {
        public String name;
        public String color;
        public DataColumn strainData;
        public DataColumn stressData;

        public ReferenceSample(JToken sampleDescription, String name)
        {
            this.name = name;
            strainData = new DataColumn(sampleDescription["strain"]);
            stressData = new DataColumn(sampleDescription["stress"]);

            // color = (String)sampleDescription["color"];
        }
    }
}
