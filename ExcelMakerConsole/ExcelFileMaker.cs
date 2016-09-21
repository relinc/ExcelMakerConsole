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

            foreach(string groupDir in Directory.GetDirectories(jobPath))
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
                    sample.column1 = new DataColumn(lines.Length-1);
                    sample.column1.dataSetInfo = datasets[0];
                    sample.column2 = new DataColumn(lines.Length-1);
                    sample.column2.dataSetInfo = datasets[1];
                    sample.column3 = new DataColumn(lines.Length-1);
                    sample.column3.dataSetInfo = datasets[2];
                    sample.column4 = new DataColumn(lines.Length-1);
                    sample.column4.dataSetInfo = datasets[3];

                    for (int i = 0; i < sample.column1.data.Length; i++)
                    {
                        string line = lines[i];
                        sample.column1.data[i] = Double.Parse(line.Split(',')[0]);
                        sample.column2.data[i] = Double.Parse(line.Split(',')[1]);
                        sample.column3.data[i] = Double.Parse(line.Split(',')[2]);
                        sample.column4.data[i] = Double.Parse(line.Split(',')[3]);
                    }

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

            String strainTitle = datasets[2].dataName;
            String stressTitle = datasets[1].dataName;
            String strainRateTitle = datasets[3].dataName;
            String timeTitle = datasets[0].dataName;

           


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
                int columnCountName = 21;
                int columnCountLabels = 21;
                int spaceName = 5;
                

                
                foreach (Sample sample in group.samples)
                {
                    double[] timeData = sample.column1.data;
                    double[] stressData = sample.column2.data;
                    double[] strainData = sample.column3.data;
                    double[] strainRateData = sample.column4.data;

                    Sheet.Cells[GetExcelColumnName(columnCountName) + "1"].Value = sample.name;
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
                    columnCountLabels++;
                    
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


                int newColumnHunter = 21;
                int trialnumber = 1;
                bool trialAddedToSummary = false;
                foreach (Sample sample in group.samples)
                {
           
                    int endIndex = sample.column1.data.Length - 1;
                    //Console.WriteLine(combinedReport.Workbook.Worksheets[group.name] + " : Worksheet");
                    //Console.WriteLine(combinedReport.Workbook.Worksheets[group.name].Cells + " : WorksheetCells");
                    //Console.WriteLine(combinedReport.Workbook.Worksheets[group.name].Cells.Count() + " : WorksheetCellsCount");
                    //Console.ReadKey();
                    var timeExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter) + "3:" + GetExcelColumnName(newColumnHunter) + (endIndex + 2).ToString()];
                    var StrainRateExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 1) + "3:" + GetExcelColumnName(newColumnHunter + 1) + (endIndex + 2).ToString()];
                    var strainExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 2) + "3:" + GetExcelColumnName(newColumnHunter + 2) + (endIndex + 2).ToString()];
                    var stressExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 3) + "3:" + GetExcelColumnName(newColumnHunter + 3) + (endIndex + 2).ToString()];
                    //var forceIncidentExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 4) + "3:" + GetExcelColumnName(newColumnHunter + 4) + (endIndex + 2).ToString()];
                    //var forceTransmissionExcel = combinedReport.Workbook.Worksheets[group.name].Cells[GetExcelColumnName(newColumnHunter + 5) + "3:" + GetExcelColumnName(newColumnHunter + 5) + (endIndex + 2).ToString()];
                    newColumnHunter += 5; //maybe should be 3

                    
                        var strainRateChartSeries = strainRateChart.Series.Add(StrainRateExcel, timeExcel);
                        var strainChartSeries = strainChart.Series.Add(strainExcel, timeExcel);
                        var stressChartSeries = stressChart.Series.Add(stressExcel, timeExcel);
                        var stressStrainSeries = stressStrainChart.Series.Add(stressExcel, strainExcel);
                        //var forceIncidentSeries = forceChart.Series.Add(forceIncidentExcel, timeExcel);
                        //var forceTransmissionSeries = forceChart.Series.Add(forceTransmissionExcel, timeExcel);
                        if ((oneTrialOnSummaryPerGroup && !trialAddedToSummary) || !oneTrialOnSummaryPerGroup)
                        {
                            var sumstrainSeries = SumstrainChart.Series.Add(strainExcel, timeExcel);
                            var sumStressSeries = SumstressChart.Series.Add(stressExcel, timeExcel);
                            var sumStrainRateSeries = SumstrainRateChart.Series.Add(StrainRateExcel, timeExcel);
                            var sumStressStrainSeries = SumstressStrainChart.Series.Add(stressExcel, strainExcel);
                            //var sumForceIncidentSeries = SumForceEquilibriumChart.Series.Add(forceIncidentExcel, timeExcel);
                            //var sumForceTransmissionSeries = SumForceEquilibriumChart.Series.Add(forceTransmissionExcel, timeExcel);


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
                    //sumForceIncidentSeries.LineColor = colors[sumColorSpot];
                    //sumForceTransmissionSeries.LineColor = "#110000";

                        strainRateChartSeries.Header = sample.name;// "Trial " + trialnumber.ToString();
                        strainChartSeries.Header = sample.name;//"Trial " + trialnumber.ToString();
                        stressChartSeries.Header = sample.name;//"Trial " + trialnumber.ToString();
                        stressStrainSeries.Header = sample.name;//"Trial " + trialnumber.ToString();
                    //forceIncidentSeries.Header = "Trial " + trialnumber.ToString() + " Incident";
                    //forceTransmissionSeries.Header = "Trial " + trialnumber.ToString() + " Transmission";

                    int colorSpot = (trialnumber - 1) % trialColors.Length;
                    int transColor = (trialnumber) % trialColors.Length;
                    strainRateChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    strainChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressChartSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    stressStrainSeries.LineColor = sample.color == null ? trialColors[colorSpot] : sample.color;
                    //forceIncidentSeries.LineColor = trialColors[colorSpot];
                    //forceTransmissionSeries.LineColor = trialColors[transColor];
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
