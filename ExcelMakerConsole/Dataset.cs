using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMakerConsole
{
    class Dataset
    {
        public String dataName; //(Stress/Displacement)
        public String dataUnits;
        public String dataType = ""; //(True/Engineering)
        

        public Dataset(string line)
        {
            if (line.Split('$').Length < 3)
            {
                Console.WriteLine("Failed!: " + line);
                return;
            }
            dataName = line.Split('$')[1];
            dataUnits = line.Split('$')[2];
            if (line.Split('$').Length >= 4)
                dataType = line.Split('$')[3];
        }
    }
}
