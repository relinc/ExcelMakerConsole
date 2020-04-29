using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelMakerConsole;
using System.IO;

namespace TestExcelMaker
{
    [TestClass]
    public class UnitTest1
    {

        

        [TestMethod]
        public void TestReferenceSample()
        {
            string solution_dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            
            ExcelFileMaker maker = new ExcelFileMaker(Path.Combine(solution_dir, "JobFiles", "OneReferenceSampleJobFile"));

            maker.exportPath = Path.Combine(solution_dir, "OneReferenceSampleJobFile.xlsx");

            maker.exportExcelFile();
        }

        [TestMethod]
        public void TestOneSample()
        {
            string solution_dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            ExcelFileMaker maker = new ExcelFileMaker(Path.Combine(solution_dir, "JobFiles", "OneSampleJobFile"));

            maker.exportPath = Path.Combine(solution_dir, "OneSampleJobFile.xlsx");

            maker.exportExcelFile();
        }

        [TestMethod]
        public void TestMultipleEverything()
        {
            string solution_dir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            ExcelFileMaker maker = new ExcelFileMaker(Path.Combine(solution_dir, "JobFiles", "MultipleEverythingJobFile"));

            maker.exportPath = Path.Combine(solution_dir, "MultipleEverythingJobFile.xlsx");

            maker.exportExcelFile();
        }
    }
}
