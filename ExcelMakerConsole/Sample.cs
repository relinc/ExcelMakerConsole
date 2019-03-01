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
        public DataColumn[] columns;
        public String color;
      
         public bool has_face_force()
        {
            return columns.Length == 6;
        } 
       
        public void read_columns(String[] lines, List<Dataset> datasets)
        {
            int column_number = lines[0].Split(',').Length;
            //allocate vector 
            columns = new DataColumn[column_number];

            for (int idx_col=0; idx_col<column_number; idx_col++)
            {
                columns[idx_col] = new DataColumn(lines.Length - 1);
                columns[idx_col].dataSetInfo = datasets[idx_col];
            }
            for (int idx_line = 0; idx_line < columns[0].data.Length; idx_line++)
            {
                string line = lines[idx_line];
                String[] line_components = line.Split(',');
                for(int idx_col=0; idx_col<column_number; idx_col++)
                {
                    columns[idx_col].data[idx_line] = Double.Parse(line_components[idx_col]);
                }
            }
        }

        internal void readParametersFile(string parametersFile)
        {
            String file = File.ReadAllText(parametersFile);
            if (file.Split('\n')[0].Split('$').Length >= 2)
                color = file.Split('\n')[0].Split('$')[1];
        }
    }
}
