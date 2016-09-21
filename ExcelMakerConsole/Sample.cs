using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelMakerConsole
{
    class Sample
    {
        public string name;
        public DataColumn column1;
        public DataColumn column2;
        public DataColumn column3;
        public DataColumn column4;
        public String color;


        internal void readParametersFile(string parametersFile)
        {
            String file = File.ReadAllText(parametersFile);
            if (file.Split('\n')[0].Split('$').Length >= 2)
                color = file.Split('\n')[0].Split('$')[1];
        }
    }
}
