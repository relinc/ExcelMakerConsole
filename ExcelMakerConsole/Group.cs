using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMakerConsole
{
    class Group
    {
        public string name;
        public List<Sample> samples = new List<Sample>();

        internal void setParametersFromString(string groupInfo)
        {
            String[] lines = groupInfo.Split('\n');
            foreach(string line in lines)
            {
                setParametersFromLine(line);
                

            }
        }

        private void setParametersFromLine(string line)
        {
            if (line.Split('$').Length < 2)
                return;
            String description = line.Split('$')[0];
            String value = line.Split('$')[1];
            if (description.Equals("Group"))
                name = value;
        }
    }
}
