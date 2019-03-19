using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace ExcelMakerConsole
{
    class Sample
    {
        public string name;
        public DataColumn strain;
        public DataColumn strainRate;
        public DataColumn stress;
        public DataColumn time;
        public DataColumn frontFaceForce = null;
        public DataColumn backFaceForce = null;
        public String color;
        public Sample(JObject sampleDescription, string sampleName)
        {
            name = sampleName;
            strain = new DataColumn(sampleDescription["strain"]);
            stress = new DataColumn(sampleDescription["stress"]);
            strainRate = new DataColumn(sampleDescription["strainRate"]);
            time = new DataColumn(sampleDescription["time"]);
            if (sampleDescription["frontFaceForce"].HasValues && sampleDescription["backFaceForce"].HasValues)
            {
                frontFaceForce = new DataColumn(sampleDescription["frontFaceForce"]);
                backFaceForce = new DataColumn(sampleDescription["backFaceForce"]);
            }
            color = (String)sampleDescription["color"];
        }
         public bool hasFaceForce()
        {
            return frontFaceForce != null && backFaceForce != null;
        } 
       
    }
}
