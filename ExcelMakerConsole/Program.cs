using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMakerConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Began program");

            //String path = "E:\\JobFile";

            ExcelFileMaker maker = new ExcelFileMaker(args[0]);

            maker.exportExcelFile();
        }
    }
}
