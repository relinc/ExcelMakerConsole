using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using Newtonsoft.Json.Linq;

namespace ExcelMakerConsole
{
    class ExcelFileMaker
    {
        String[] summaryColors = { "E90000", "EC8600", "000ED4", "7A378B", "008080", "FFD700", "FF8247", "8E8E38", "1B85B8", "5A5255", "559E83", "AE5A41", "C3Cb71" };
        String[] trialColors = { "3ba0c1", "ff8f20", "909090", "00bf00", "800080", "007580", "df0000" };
        private String jobFilePath;
        private bool makeSummarypage;
        private bool oneTrialOnSummaryPerGroup;
        private String exportPath;
        private int version;
        private List<Group> groups = new List<Group>();
        private void setParameters(JObject jobDescription)
        {
            makeSummarypage = (bool)jobDescription["Summary_Page"];
            version = (int)jobDescription["JSON_Version"];
            exportPath = (string)jobDescription["Export_Location"];
        }
        public ExcelFileMaker(String jobPath)
        {
            //jobPath = jobfile directory
            JObject jobDescription = JObject.Parse(File.ReadAllText(jobPath + "\\" + "Description.json"));
            setParameters(jobDescription);
            Console.WriteLine("Job Parameters: " + jobDescription.ToString());
            Console.WriteLine("Done");


            Console.WriteLine("Make summmary page:" + makeSummarypage);
            Console.WriteLine("Export path: " + exportPath);
            Console.WriteLine("Version: " + version);
            foreach (JToken groupJSONToken in (JArray)jobDescription["groups"])
            {
                string groupName = (string)groupJSONToken;
                Group group = new Group();
                group.name = groupName;
                JArray groupDescription = JArray.Parse(File.ReadAllText(jobPath + "\\" + group.name + ".json"));
                foreach (JToken sampleToken in groupDescription)
                {
                    
                    string sampleName = (string)sampleToken;
                    JObject sampleDescription = JObject.Parse(File.ReadAllText(jobPath + "\\" + groupName + "\\" + sampleName + ".json"));
                    Sample sample = new Sample(sampleDescription, sampleName);
                    group.samples.Add(sample);
                }
                groups.Add(group);
            }
             Console.WriteLine("Done building sample structure");
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
   
        internal void exportExcelFile()
        {
            //double timeScale = getTimeScaleFromString(scale);
            //bool summaryPage = summaryPage;
            string combindReportName = exportPath.Replace("\r","");
            Sample exemplarSample = groups[0].samples[0];
            Dataset strainDataset = exemplarSample.strain.dataSetInfo;
            Dataset stressDataset = exemplarSample.stress.dataSetInfo;
            Dataset strainRateDataset = exemplarSample.strainRate.dataSetInfo;
            Dataset timeDataset = exemplarSample.time.dataSetInfo;
            Dataset frontFaceForceDataset = null;
            Dataset backFaceForceDataset = null;
            bool hasFaceForces = false;
            if (exemplarSample.hasFaceForce())
            {
                hasFaceForces = true;
                frontFaceForceDataset = exemplarSample.frontFaceForce.dataSetInfo;
                backFaceForceDataset = exemplarSample.backFaceForce.dataSetInfo;
            }
                FileInfo newFile = new FileInfo(combindReportName);
            if (newFile.Exists)
            {
                try
                {
                    newFile.Delete();  // ensures we create a new workbook
                    newFile = new FileInfo(combindReportName);
                }
                catch
                {
                    Console.WriteLine("Close File to Save Over");
                    //MessageBox.Show("Close File to save over");
                }
            }
            ExcelPackage combinedReport = new ExcelPackage(newFile);
            String strainHeaderUnits = (strainDataset.dataType == "" ? "" : strainDataset.dataType + " ") + strainDataset.dataName + " " + strainDataset.dataUnits;
            String stressHeaderUnits = (stressDataset.dataType == "" ? "" : stressDataset.dataType + " ") + stressDataset.dataName + " " + stressDataset.dataUnits;
            String strainRateHeaderUnits = (strainRateDataset.dataType == "" ? "" : strainRateDataset.dataType + " ") + strainRateDataset.dataName + " " + strainRateDataset.dataUnits;
            String timeHeaderUnits = timeDataset.dataName + " " + timeDataset.dataUnits;//"Time (" + scale + ")";
            String frontFaceForceUnits = "";
            String backFaceForceUnits = "";
            
            if (hasFaceForces)
            {
                 frontFaceForceUnits = (frontFaceForceDataset.dataType == "" ? "" : frontFaceForceDataset.dataType + " ") + frontFaceForceDataset.dataName + " " + frontFaceForceDataset.dataUnits;
                 backFaceForceUnits = (backFaceForceDataset.dataType == "" ? "" : backFaceForceDataset.dataType + " ") + backFaceForceDataset.dataName + " " + backFaceForceDataset.dataUnits;
            }

            String strainTitle = strainDataset.dataName;
            String stressTitle = stressDataset.dataName;
            String strainRateTitle = strainRateDataset.dataName;
            String timeTitle = timeDataset.dataName;
            String frontFaceForceTitle = "";
            String backFaceForceTitle = "";
            if(hasFaceForces)
            {
                frontFaceForceTitle = frontFaceForceDataset.dataName;
                backFaceForceTitle = backFaceForceDataset.dataName;

            }



            var SummarySheet = combinedReport.Workbook.Worksheets.Add("Summary");
            
            var SumstrainRateChart = SummarySheet.Drawings.AddChart(strainRateHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
            SumstrainRateChart.SetPosition(400, 0);
            SumstrainRateChart.SetSize(600, 400);
            SumstrainRateChart.XAxis.Title.Text = timeHeaderUnits;
            SumstrainRateChart.YAxis.Title.Text = strainRateHeaderUnits;
            SumstrainRateChart.Title.Text = strainRateTitle + " vs " + timeTitle;
            SumstrainRateChart.YAxis.MinValue = 0;
            SumstrainRateChart.XAxis.MinValue = 0;

            var SumstrainChart = SummarySheet.Drawings.AddChart(strainHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
            SumstrainChart.SetPosition(0, 0);
            SumstrainChart.SetSize(600, 400);
            SumstrainChart.XAxis.Title.Text = timeHeaderUnits;
            SumstrainChart.YAxis.Title.Text = strainHeaderUnits;
            SumstrainChart.Title.Text = strainTitle + " vs " + timeTitle;
            SumstrainChart.YAxis.MinValue = 0;
            SumstrainChart.XAxis.MinValue = 0;

            var SumstressChart = SummarySheet.Drawings.AddChart(stressHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
            SumstressChart.SetPosition(0, 600);
            SumstressChart.SetSize(600, 400);
            SumstressChart.XAxis.Title.Text = timeHeaderUnits;
            SumstressChart.YAxis.Title.Text = stressHeaderUnits;
            SumstressChart.Title.Text = stressTitle + " vs " + timeTitle;
            SumstressChart.YAxis.MinValue = 0;
            SumstressChart.XAxis.MinValue = 0;

            var SumstressStrainChart = SummarySheet.Drawings.AddChart(stressHeaderUnits + " vs " + strainHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
            SumstressStrainChart.SetPosition(400, 600);
            SumstressStrainChart.SetSize(600, 400);
            SumstressStrainChart.XAxis.Title.Text = strainHeaderUnits;
            SumstressStrainChart.YAxis.Title.Text = stressHeaderUnits;
            SumstressStrainChart.Title.Text = stressTitle + " vs " + strainTitle;
            SumstressStrainChart.YAxis.MinValue = 0;
            SumstressStrainChart.XAxis.MinValue = 0;

            int idx = 0;
            foreach(Group group in groups)
            {

                var Sheet = combinedReport.Workbook.Worksheets.Add(group.name);
                FileInfo placeHolderFile = new FileInfo("tmp");
                ExcelPackage placeHolderExcel = new ExcelPackage(placeHolderFile);
                var fakeSheet = placeHolderExcel.Workbook.Worksheets.Add("placeholder");
                int columnCountName = 21;
                int columnCountLabels = 21;
                int spaceName = 5;
                

                
                foreach (Sample sample in group.samples)
                {
                    double[] timeData = sample.time.data;
                    double[] stressData = sample.stress.data;
                    double[] strainData = sample.strain.data;
                    double[] strainRateData = sample.strainRate.data;
                   

                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "1"].Value = sample.name;
                    columnCountName += spaceName;
                    
                    //time
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = timeHeaderUnits;
                    for (int i = 0; i < timeData.Length; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = timeData[i];
                    }
                    columnCountLabels++;
                    //Strain Rate
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = strainRateHeaderUnits;
                    for (int i = 0; i < timeData.Length; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = strainRateData[i];
                    }
                    columnCountLabels++;
                    //Strain
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = strainHeaderUnits;
                    for (int i = 0; i < timeData.Length; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = strainData[i];
                    }
                    columnCountLabels++;
                    //Stress
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = stressHeaderUnits;
                    for (int i = 0; i < timeData.Length; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = stressData[i];
                    }
                    columnCountLabels++;
                    if (sample.hasFaceForce())
                    {
                        double[] frontFaceData = sample.frontFaceForce.data;
                        double[] backFaceData = sample.backFaceForce.data;

                        
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = frontFaceForceUnits;
                        for (int i = 0; i < timeData.Length; i++)
                        {
                            Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = frontFaceData[i];
                        }
                        columnCountLabels++;
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = backFaceForceUnits;
                        for (int i = 0; i < timeData.Length; i++)
                        {
                            Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = backFaceData[i];
                        }
                        columnCountLabels++;
                        columnCountLabels++;
                    }
                    else
                    {
                        columnCountLabels+=3;

                    }
                }

                var strainChart = Sheet.Drawings.AddChart(strainHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                strainChart.SetPosition(0, 0);
                strainChart.SetSize(600, 400);

                strainChart.XAxis.Title.Text = timeHeaderUnits;
                strainChart.YAxis.Title.Text = strainHeaderUnits;
                strainChart.Title.Text = strainTitle + " vs " + timeTitle;
                strainChart.YAxis.MinValue = 0;
                strainChart.XAxis.MinValue = 0;

                var strainRateChart = Sheet.Drawings.AddChart(strainRateHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                strainRateChart.SetPosition(400, 0);
                strainRateChart.SetSize(600, 400);

                strainRateChart.XAxis.Title.Text = timeHeaderUnits;
                strainRateChart.YAxis.Title.Text = strainRateHeaderUnits;
                strainRateChart.Title.Text = strainRateTitle + " vs " + timeTitle;
                strainRateChart.YAxis.MinValue = 0;
                strainRateChart.XAxis.MinValue = 0;

                var stressChart = Sheet.Drawings.AddChart(stressHeaderUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                 stressChart.SetPosition(0, 600);
                stressChart.SetSize(600, 400);

                stressChart.XAxis.Title.Text = timeHeaderUnits;
                stressChart.YAxis.Title.Text = stressHeaderUnits;
                stressChart.Title.Text = stressTitle + " vs " + timeTitle;
                stressChart.YAxis.MinValue = 0;
                stressChart.XAxis.MinValue = 0;

                var stressStrainChart = Sheet.Drawings.AddChart(stressHeaderUnits + " vs " + strainHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                  stressStrainChart.SetPosition(400, 600);
                stressStrainChart.SetSize(600, 400);

                stressStrainChart.XAxis.Title.Text = strainHeaderUnits;
                stressStrainChart.YAxis.Title.Text = stressHeaderUnits;
                stressStrainChart.Title.Text = stressTitle + " vs " + strainTitle;
                stressStrainChart.YAxis.MinValue = 0;
                stressStrainChart.XAxis.MinValue = 0;
                ExcelChart frontFaceForceChart = null;
                ExcelChart backFaceForceChart = null;

                if (hasFaceForces)
                {
                    frontFaceForceChart = Sheet.Drawings.AddChart("front " + frontFaceForceUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                    backFaceForceChart = Sheet.Drawings.AddChart("back " + backFaceForceUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);


                    frontFaceForceChart.SetPosition(800, 0);
                    frontFaceForceChart.SetSize(600, 400);
                    frontFaceForceChart.XAxis.Title.Text = timeHeaderUnits;
                    frontFaceForceChart.YAxis.Title.Text = frontFaceForceUnits;
                    frontFaceForceChart.Title.Text = frontFaceForceTitle + " vs " + timeTitle;
                    frontFaceForceChart.YAxis.MinValue = 0;
                    frontFaceForceChart.XAxis.MinValue = 0;


                    backFaceForceChart.SetPosition(800, 600);
                    backFaceForceChart.SetSize(600, 400);
                    backFaceForceChart.XAxis.Title.Text = timeHeaderUnits;
                    backFaceForceChart.YAxis.Title.Text = backFaceForceUnits;
                    backFaceForceChart.Title.Text = backFaceForceTitle + " vs " + timeTitle;
                    backFaceForceChart.YAxis.MinValue = 0;
                    backFaceForceChart.XAxis.MinValue = 0;
                }
                int newColumnHunter = 21;
                int trialnumber = 1;
                bool trialAddedToSummary = false;
                foreach (Sample sample in group.samples)
                {
           
                    int endIndex = sample.time.data.Length - 1;
                    var timeExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter) + "3:" + GetExcelColumnName(newColumnHunter) + (endIndex + 2).ToString()];
                    var StrainRateExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter +1) + "3:" + GetExcelColumnName(newColumnHunter + 1) + (endIndex + 2).ToString()];
                    var strainExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 2) + "3:" + GetExcelColumnName(newColumnHunter + 2) + (endIndex + 2).ToString()];
                    var stressExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 3) + "3:" + GetExcelColumnName(newColumnHunter + 3) + (endIndex + 2).ToString()];
                    var forceIncidentExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 4) + "3:" + GetExcelColumnName(newColumnHunter + 4) + (endIndex + 2).ToString()];
                    var forceTransmissionExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 5) + "3:" + GetExcelColumnName(newColumnHunter + 5) + (endIndex + 2).ToString()];

                    newColumnHunter += 7;
                   
                    
                        var strainRateChartSeries = strainRateChart.Series.Add(StrainRateExcel, timeExcel);
                        var strainChartSeries = strainChart.Series.Add(strainExcel, timeExcel);
                        var stressChartSeries = stressChart.Series.Add(stressExcel, timeExcel);
                        var stressStrainSeries = stressStrainChart.Series.Add(stressExcel, strainExcel);
                    //initialize the face force charts
                    ExcelChartSerie forceIncidentSeries= null;  //always_fake.Series.Add(stressExcel, strainExcel);
                    ExcelChartSerie forceTransmissionSeries=null; //always_fake.Series.Add(stressExcel, strainExcel);
                    if (sample.hasFaceForce()) {
                            forceIncidentSeries = frontFaceForceChart.Series.Add(forceIncidentExcel, timeExcel);
                            forceTransmissionSeries = backFaceForceChart.Series.Add(forceTransmissionExcel, timeExcel);
                         }
                                                   
                        if ((oneTrialOnSummaryPerGroup && !trialAddedToSummary) || !oneTrialOnSummaryPerGroup)
                        {
                            var sumstrainSeries = SumstrainChart.Series.Add(strainExcel, timeExcel);
                            var sumStressSeries = SumstressChart.Series.Add(stressExcel, timeExcel);
                            var sumStrainRateSeries = SumstrainRateChart.Series.Add(StrainRateExcel, timeExcel);
                            var sumStressStrainSeries = SumstressStrainChart.Series.Add(stressExcel, strainExcel);
                            


                            sumstrainSeries.Header = group.name;
                            sumStressSeries.Header = group.name;
                            sumStrainRateSeries.Header = group.name;
                            sumStressStrainSeries.Header = group.name;
                            



                            int sumColorSpot = (idx) % summaryColors.Length;
                            sumstrainSeries.LineColor = summaryColors[sumColorSpot];
                            sumStressSeries.LineColor = summaryColors[sumColorSpot];
                            sumStrainRateSeries.LineColor = summaryColors[sumColorSpot];
                            sumStressStrainSeries.LineColor = summaryColors[sumColorSpot];
                            
                        }
                        strainRateChartSeries.Header = sample.name;
                        strainChartSeries.Header = sample.name;
                        stressChartSeries.Header = sample.name;
                        stressStrainSeries.Header = sample.name;


                    int colorSpot = (trialnumber - 1) % trialColors.Length;
                    int transColor = (trialnumber) % trialColors.Length;
                    strainRateChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    strainChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressStrainSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    if(sample.hasFaceForce()){
                        forceIncidentSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                        forceTransmissionSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                        forceIncidentSeries.Header = sample.name;
                        forceTransmissionSeries.Header = sample.name;
                    }
                    trialnumber++;
                    trialAddedToSummary = true;
                }
             
                idx++;
            }

            if (!makeSummarypage)
            {
             
                combinedReport.Workbook.Worksheets.Delete(combinedReport.Workbook.Worksheets["Summary"]);
            }
           
            combinedReport.Save();
            


        
        }
    }
}
