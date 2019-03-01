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
        private List<Dataset> datasets = new List<Dataset>();
        private List<Group> groups = new List<Group>();
        public ExcelFileMaker(String jobPath)
        {
            //jobPath = jobfile directory
            String jobParameters = File.ReadAllText(jobPath + "\\" + "Parameters.txt");
            Console.WriteLine("Job Parameters: " + jobParameters);
            Console.WriteLine("Done");

            //String infoSection = "";
           // jobFilePath = jobPath;
           // string readText = File.ReadAllText(jobPath);
           // String[] sections = readText.Split(new string[] { "%#%@@!!!!" }, StringSplitOptions.None);
           //String infoSection = sections[0];
            setMakerParameters(jobParameters);

            Console.WriteLine("Make summmary page:" + makeSummarypage);
            Console.WriteLine("Export path: " + exportPath);
            Console.WriteLine("Version: " + version);
            Console.WriteLine("Num datasets: " + datasets.Count);
            foreach (string groupDir in Directory.GetDirectories(jobPath))
            {
                String groupName = Path.GetFileName(groupDir);
                Group group = new Group();
                group.name = groupName;
                foreach(string sampleDir in Directory.GetDirectories(groupDir))
                {
                    Sample sample = new Sample();
                    sample.name = Path.GetFileName(sampleDir);

                    String parametersFile = sampleDir + "\\" + "Parameters.txt";
                    if (File.Exists(parametersFile))
                        sample.readParametersFile(parametersFile);
                    String dataFile = File.ReadAllText(sampleDir + "\\" + "Data.txt");
                    String[] lines = dataFile.Split('\n');
                    sample.read_columns(lines, datasets);
                    group.samples.Add(sample);
                }
                groups.Add(group);
            }
             Console.WriteLine("Done building sample structure");
        }

        private void setMakerParameters(string infoSection)
        {
            foreach (String line in infoSection.Split('\n'))
                setParameterFromLine(line);
        }

        private void setParameterFromLine(string line)
        {
            if (line.Split('$').Length < 2)
                return;
            String description = line.Split('$')[0];
            String val = line.Split('$')[1];
            if (description.Equals("Version"))
                version = int.Parse(val);
            else if (description.Equals("Export Location"))
                exportPath = val;
            else if (description.Equals("Summary Page"))
                makeSummarypage = bool.Parse(val);
            else if (description.StartsWith("Dataset"))
                datasets.Add(new Dataset(line));
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
        private bool datasets_have_face_force(List<Dataset> datasets)
        {
            return datasets.Count == 6;
        }
        internal void exportExcelFile()
        {
            
            //double timeScale = getTimeScaleFromString(scale);
            //bool summaryPage = summaryPage;
            string combindReportName = exportPath.Replace("\r","");
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
            String strainHeaderUnits = (datasets[2].dataType == "" ? "" : datasets[2].dataType + " ") + datasets[2].dataName + " " + datasets[2].dataUnits;
            String stressHeaderUnits = (datasets[1].dataType == "" ? "" : datasets[1].dataType + " ") + datasets[1].dataName + " " + datasets[1].dataUnits;
            String strainRateHeaderUnits = (datasets[3].dataType == "" ? "" : datasets[3].dataType + " ") + datasets[3].dataName + " " + datasets[3].dataUnits;
            String timeHeaderUnits = datasets[0].dataName + " " + datasets[0].dataUnits;//"Time (" + scale + ")";
            String frontFaceForceUnits = "";
            String backFaceForceUnits = "";
            if (datasets_have_face_force(datasets))
            {
                 frontFaceForceUnits = (datasets[4].dataType == "" ? "" : datasets[4].dataType + " ") + datasets[4].dataName + " " + datasets[4].dataUnits;
                 backFaceForceUnits = (datasets[5].dataType == "" ? "" : datasets[5].dataType + " ") + datasets[5].dataName + " " + datasets[5].dataUnits;
            }

            String strainTitle = datasets[2].dataName;
            String stressTitle = datasets[1].dataName;
            String strainRateTitle = datasets[3].dataName;
            String timeTitle = datasets[0].dataName;
            String frontFaceForceTitle = "";
            String backFaceForceTitle = "";
            if(datasets_have_face_force(datasets))
            {
                frontFaceForceTitle = datasets[4].dataName;
                backFaceForceTitle = datasets[5].dataName;

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
                FileInfo place_holder_file = new FileInfo("tmp");
                ExcelPackage place_holder_excel = new ExcelPackage(place_holder_file);
                var fake_sheet = place_holder_excel.Workbook.Worksheets.Add("placeholder");
                int columnCountName = 21;
                int columnCountLabels = 21;
                int spaceName = 5;
                

                
                foreach (Sample sample in group.samples)
                {
                    double[] timeData = sample.columns[0].data;
                    double[] stressData = sample.columns[1].data;
                    double[] strainData = sample.columns[2].data;
                    double[] strainRateData = sample.columns[3].data;
                    double[] frontFaceData = new double[0];
                    double[] backFaceData = new double[0];
                    if (sample.has_face_force())
                    {
                        frontFaceData = sample.columns[4].data;
                        backFaceData = sample.columns[5].data;
                    }

                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "1"].Value = sample.name;
                    columnCountName += spaceName;
                    
                    //time
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = timeHeaderUnits;
                    for (int i = 0; i < timeData.Length-1; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = timeData[i];
                    }
                    columnCountLabels++;
                    //Strain Rate
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = strainRateHeaderUnits;
                    for (int i = 0; i < timeData.Length-1; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = strainRateData[i];
                    }
                    columnCountLabels++;
                    //Strain
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = strainHeaderUnits;
                    for (int i = 0; i < timeData.Length-1; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = strainData[i];
                    }
                    columnCountLabels++;
                    //Stress
                    Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = stressHeaderUnits;
                    for (int i = 0; i < timeData.Length-1; i++)
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = stressData[i];
                    }
                    columnCountLabels++;
                    if (sample.has_face_force())
                    {
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = frontFaceForceUnits;
                        for (int i = 0; i < timeData.Length-1; i++)
                        {
                            Sheet.Cells[GetExcelColumnName(columnCountLabels) + (i + 3).ToString()].Value = frontFaceData[i];
                        }
                        columnCountLabels++;
                        Sheet.Cells[GetExcelColumnName(columnCountLabels) + "2"].Value = backFaceForceUnits;
                        for (int i = 0; i < timeData.Length-1; i++)
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

                var frontFaceForceChart = fake_sheet.Drawings.AddChart("placeholder", eChartType.XYScatterSmoothNoMarkers);
                var backFaceForceChart = fake_sheet.Drawings.AddChart("placeholder2", eChartType.XYScatterSmoothNoMarkers);
                var always_fake = fake_sheet.Drawings.AddChart("placeholder3", eChartType.XYScatterSmoothNoMarkers);

                if (datasets_have_face_force(datasets))
                {
                    frontFaceForceChart = Sheet.Drawings.AddChart(frontFaceForceUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);
                    backFaceForceChart = Sheet.Drawings.AddChart(backFaceForceUnits + " vs " + timeHeaderUnits, eChartType.XYScatterSmoothNoMarkers);


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
           
                    int endIndex = sample.columns[0].data.Length - 1;
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

                    ExcelChartSerie forceIncidentSeries = always_fake.Series.Add(stressExcel, strainExcel);
                    ExcelChartSerie forceTransmissionSeries = always_fake.Series.Add(stressExcel, strainExcel);
                    if (sample.has_face_force()) {
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
                        strainRateChartSeries.Header = sample.name;// "Trial " + trialnumber.ToString();
                        strainChartSeries.Header = sample.name;//"Trial " + trialnumber.ToString();
                        stressChartSeries.Header = sample.name;//"Trial " + trialnumber.ToString();
                        stressStrainSeries.Header = sample.name;//"Trial " + trialnumber.ToString();


                    int colorSpot = (trialnumber - 1) % trialColors.Length;
                    int transColor = (trialnumber) % trialColors.Length;
                    strainRateChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    strainChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressStrainSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    if(sample.has_face_force()){
                        forceIncidentSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                        forceTransmissionSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                        forceIncidentSeries.Header = sample.name;// "Trial " + trialnumber.ToString() + " Incident";
                        forceTransmissionSeries.Header = sample.name;// "Trial " + trialnumber.ToString() + " Transmission";
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
