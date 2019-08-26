using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using ARC_Head_Counts;

namespace BackedExcelFunctions
{
    class _Excel
    {
        _Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb;
        Worksheet ws;
        Range range;
        string filePath = @"";
        int currentSheet;
        double[][] jaggedArray;

        public void ShowGraph(int comboboxGraphindex, int comboBoxIndex, int fileSelected)
        {

            //Graphs are most likely not working because the file isn't being properly created, thus not being able to be displayed.

            ws = wb.Worksheets[comboBoxIndex];
            double r = ws.Cells[22, 21].Value;
            string endRange = "";

            if (fileSelected == 0)
            {
                endRange = "R";
            }
            else if (fileSelected == 1)
            {
                endRange = "H";
            }
            else if (fileSelected == 2)
            {
                endRange = "W";
            }

            range = ws.Range["B1", endRange + r.ToString()];
            object misValue = System.Reflection.Missing.Value;
            string pathToGraph = "";
            Random random = new Random();
            string newFolderDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts";
            xlApp.DisplayAlerts = false;
            Directory.CreateDirectory(newFolderDirectory);

            switch (comboboxGraphindex)
            {
                case 1:
                    int rand_1 = random.Next(int.MinValue, int.MaxValue);
                    ChartObjects xlCharts_1 = (ChartObjects)ws.ChartObjects(Type.Missing);
                    ChartObject myChart_1 = (ChartObject)xlCharts_1.Add(0, 400, 925, 575);
                    Chart chartPage_1 = myChart_1.Chart;
                    chartPage_1.SetSourceData(range);
                    chartPage_1.ChartType = XlChartType.xl3DColumn;
                    chartPage_1.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
                    myChart_1.Activate();
                    chartPage_1.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_1.ToString() +".PNG", "PNG", misValue);
                    pathToGraph = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_1.ToString() + ".PNG";
                    Graph_Form graph_1 = new Graph_Form();
                    graph_1.pictureBoxGraphs.Image = Image.FromFile(pathToGraph);
                    graph_1.Show();
                    xlApp.DisplayAlerts = false;
                    break;
                case 2:
                    int rand_2 = random.Next(int.MinValue, int.MaxValue);
                    ChartObjects xlCharts_2 = (ChartObjects)ws.ChartObjects(Type.Missing);
                    ChartObject myChart_2 = (ChartObject)xlCharts_2.Add(0, 400, 925, 578);
                    Chart chartPage_2 = myChart_2.Chart;
                    chartPage_2.SetSourceData(range);
                    chartPage_2.ChartType = XlChartType.xl3DAreaStacked;
                    chartPage_2.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
                    myChart_2.Activate();
                    chartPage_2.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_2.ToString() + ".PNG", "PNG", misValue);
                    pathToGraph = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_2.ToString() + ".PNG";
                    Graph_Form graph_2 = new Graph_Form();
                    graph_2.pictureBoxGraphs.Image = Image.FromFile(pathToGraph);
                    graph_2.Show();
                    xlApp.DisplayAlerts = false;
                    break;
                case 3:
                    
                    int rand_3 = random.Next(int.MinValue, int.MaxValue);
                    ChartObjects xlCharts_3 = (ChartObjects)ws.ChartObjects(Type.Missing);
                    ChartObject myChart_3 = (ChartObject)xlCharts_3.Add(0, 400, 925, 578);
                    Chart chartPage_3 = myChart_3.Chart;
                    chartPage_3.SetSourceData(range);
                    chartPage_3.ChartType = XlChartType.xlColumnClustered;
                    chartPage_3.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
                    myChart_3.Activate();
                    chartPage_3.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_3.ToString() + ".PNG", "PNG", misValue);
                    pathToGraph = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_3.ToString() + ".PNG";
                    Graph_Form graph_3 = new Graph_Form();
                    graph_3.pictureBoxGraphs.Image = Image.FromFile(pathToGraph);
                    graph_3.Show();
                    xlApp.DisplayAlerts = false;
                    break;
                case 4:
                    int rand_4 = random.Next(int.MinValue, int.MaxValue);
                    ChartObjects xlCharts_4 = (ChartObjects)ws.ChartObjects(Type.Missing);
                    ChartObject myChart_4 = (ChartObject)xlCharts_4.Add(0, 400, 925, 578);
                    Chart chartPage_4 = myChart_4.Chart;
                    chartPage_4.SetSourceData(range);
                    chartPage_4.ChartType = XlChartType.xlColumnStacked;
                    chartPage_4.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
                    myChart_4.Activate();
                    chartPage_4.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_4.ToString() + ".PNG", "PNG", misValue);
                    pathToGraph = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_4.ToString() + ".PNG";
                    Graph_Form graph_4 = new Graph_Form();
                    graph_4.pictureBoxGraphs.Image = Image.FromFile(pathToGraph);
                    graph_4.Show();
                    xlApp.DisplayAlerts = false;
                    break;
                case 5:
                    int rand_5 = random.Next(int.MinValue, int.MaxValue);
                    ChartObjects xlCharts_5 = (ChartObjects)ws.ChartObjects(Type.Missing);
                    ChartObject myChart_5 = (ChartObject)xlCharts_5.Add(0, 400, 925, 578);
                    Chart chartPage_5= myChart_5.Chart;
                    chartPage_5.SetSourceData(range);
                    chartPage_5.ChartType = XlChartType.xlLineMarkers;
                    chartPage_5.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
                    myChart_5.Activate();
                    chartPage_5.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_5.ToString() + ".PNG", "PNG", misValue);
                    pathToGraph = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + rand_5.ToString() + ".PNG";
                    Graph_Form graph_5 = new Graph_Form();
                    graph_5.pictureBoxGraphs.Image = Image.FromFile(pathToGraph);
                    graph_5.Show();
                    xlApp.DisplayAlerts = false;
                    break;
            }
            xlApp.DisplayAlerts = false;
            //wb.Save();
            //Marshal.ReleaseComObject(wb);
            //return pathToGraph;
        }
        public void TestGraph()
        {
            xlApp.DisplayAlerts = false;
            string newFolderDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts";
            Directory.CreateDirectory(newFolderDirectory);
            object misValue = System.Reflection.Missing.Value;
            range = ws.Range["B1", "R14"];
            ChartObjects xlCharts = (ChartObjects)ws.ChartObjects(Type.Missing);
            ChartObject myChart = xlCharts.Add(0, 400, 1000, 600);
            Chart chartPage = myChart.Chart;
            chartPage.ChartType = XlChartType.xl3DColumn;
            chartPage.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
            myChart.Activate();
            chartPage.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - System Charts\\Chart" + ".PNG", "PNG", misValue);
        }
        public void Graph(int bottomRange_w, int fileSelected)
        {
            //!!!!!!!!!!!!************* Number of charts can be made, not just one *************!!!!!!!!!!!!

            object misValue = System.Reflection.Missing.Value;
            string endColumnRange = "";

            if (fileSelected == 0)
            {
                endColumnRange = "R";
            }
            else if (fileSelected == 1)
            {
                endColumnRange = "H";
            }
            else if (fileSelected == 2)
            {
                endColumnRange = "W";
            }

            range = ws.Range["B1", endColumnRange + bottomRange_w.ToString()];
            ChartObjects xlCharts = (ChartObjects)ws.ChartObjects(Type.Missing);
            //                                      Position X, Position Y, Width, Height
            ChartObject myChart = (ChartObject)xlCharts.Add(0, 400, 1000, 600);
            Chart chartPage = myChart.Chart;
            chartPage.SetSourceData(range);
            chartPage.ChartType = XlChartType.xl3DColumn;
            chartPage.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
            //chartPage.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Chart.PNG", "PNG", misValue);

            ChartObject myChart_1 = (ChartObject)xlCharts.Add(1000, 400, 1000, 600);
            Chart chartPage_1 = myChart_1.Chart;
            chartPage_1.SetSourceData(range);
            chartPage_1.ChartType = XlChartType.xlAreaStacked;
            chartPage_1.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");

            ChartObject myChart_2 = (ChartObject)xlCharts.Add(0, 1000, 1000, 600);
            Chart chartPage_2 = myChart_2.Chart;
            chartPage_2.SetSourceData(range);
            chartPage_2.ChartType = XlChartType.xlColumnClustered;
            chartPage_2.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");

            ChartObject myChart_3 = (ChartObject)xlCharts.Add(1000, 1000, 1000, 600);
            Chart chartPage_3 = myChart_3.Chart;
            chartPage_3.SetSourceData(range);
            chartPage_3.ChartType = XlChartType.xlColumnStacked;
            chartPage_3.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");

            ChartObject myChart_4 = (ChartObject)xlCharts.Add(0, 1600, 1000, 600);
            Chart chartPage_4 = myChart_4.Chart;
            chartPage_4.SetSourceData(range);
            chartPage_4.ChartType = XlChartType.xlLineMarkers;
            chartPage_4.ChartWizard(Source: range, Title: "Head Count Data by Hour", CategoryTitle: "Time and Date", ValueTitle: "Number of occupants");
            wb.Save();
        }
        public void GenerateSummaryGraph(string[] times, double[][] roomsData ,int numberOfDays, int sheet, int fileSelected)
        {
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Worksheet)xlApp.Worksheets[sheet];
            ws = wb.Worksheets[sheet];
            worksheet.Name = "Summary";

            int column = 0;

            if (fileSelected == 0)
            {
                column = 20;

                ws.Cells[1, 1].Value = "Time";
                ws.Cells[1, 2].Value = "First Floor Lobby";
                ws.Cells[1, 3].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 4].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 5].Value = "ARC 120";
                ws.Cells[1, 6].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 7].Value = "ARC 130 (Student Leaders Office)";
                ws.Cells[1, 8].Value = "ARC-131 A (Small Metting Room)";
                ws.Cells[1, 9].Value = "ARC-131 B (Small Meeting Room)";
                ws.Cells[1, 10].Value = "ARC-141 A (Small Meeting Room)";
                ws.Cells[1, 11].Value = "ARC-141 B (Small Metting Room)";
                ws.Cells[1, 12].Value = "ARC- 135 (Work Room)";
                ws.Cells[1, 13].Value = "ARC- 210 (Meeting Room)";
                ws.Cells[1, 14].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 15].Value = "Sports Field";
                ws.Cells[1, 16].Value = "Volleyball/Basketball/Tennis (combined)";

                ws.Cells[1, 20].Value = "Time";
                ws.Cells[1, 21].Value = "First Floor Lobby";
                ws.Cells[1, 22].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 23].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 24].Value = "ARC 120";
                ws.Cells[1, 25].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 26].Value = "ARC 130 (Student Leaders Office)";
                ws.Cells[1, 27].Value = "ARC-131 A (Small Metting Room)";
                ws.Cells[1, 28].Value = "ARC-131 B (Small Meeting Room)";
                ws.Cells[1, 29].Value = "ARC-141 A (Small Meeting Room)";
                ws.Cells[1, 30].Value = "ARC-141 B (Small Metting Room)";
                ws.Cells[1, 31].Value = "ARC- 135 (Work Room)";
                ws.Cells[1, 32].Value = "ARC- 210 (Meeting Room)";
                ws.Cells[1, 33].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 34].Value = "Sports Field";
                ws.Cells[1, 35].Value = "Volleyball/Basketball/Tennis (combined)";
            }

            else if (fileSelected == 1)
            {
                column = 8;

                ws.Cells[1, 1].Value = "Time";
                ws.Cells[1, 2].Value = "Lower Level Lobby";
                ws.Cells[1, 3].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 4].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 5].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 6].Value = "Fitness Studio";

                ws.Cells[1, 8].Value = "Time";
                ws.Cells[1, 9].Value = "Lower Level Lobby";
                ws.Cells[1, 10].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 11].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 12].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 13].Value = "Fitness Studio";
            }
            else if (fileSelected == 2)
            {
                column = 23;

                ws.Cells[1, 1].Value = "Time";
                ws.Cells[1, 2].Value = "Lower Level Lobby";
                ws.Cells[1, 3].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 4].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 5].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 6].Value = "Fitness Studio";
                ws.Cells[1, 7].Value = "First Floor Lobby";
                ws.Cells[1, 8].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 9].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 10].Value = "ARC 120";
                ws.Cells[1, 11].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 12].Value = "ARC 130 (Student Leaders Office)";
                ws.Cells[1, 13].Value = "ARC-131 A (Small Metting Room)";
                ws.Cells[1, 14].Value = "ARC-131 B (Small Meeting Room)";
                ws.Cells[1, 15].Value = "ARC-141 A (Small Meeting Room)";
                ws.Cells[1, 16].Value = "ARC-141 B (Small Metting Room)";
                ws.Cells[1, 17].Value = "ARC- 135 (Work Room)";
                ws.Cells[1, 18].Value = "ARC- 210 (Meeting Room)";
                ws.Cells[1, 19].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 20].Value = "Sports Field";
                ws.Cells[1, 21].Value = "Volleyball/Basketball/Tennis (combined)";

                ws.Cells[1, 23].Value = "Time";
                ws.Cells[1, 24].Value = "Lower Level Lobby";
                ws.Cells[1, 25].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 26].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 27].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 28].Value = "Fitness Studio";
                ws.Cells[1, 29].Value = "First Floor Lobby";
                ws.Cells[1, 30].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 31].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 32].Value = "ARC 120";
                ws.Cells[1, 33].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 34].Value = "ARC 130 (Student Leaders Office)";
                ws.Cells[1, 35].Value = "ARC-131 A (Small Metting Room)";
                ws.Cells[1, 36].Value = "ARC-131 B (Small Meeting Room)";
                ws.Cells[1, 37].Value = "ARC-141 A (Small Meeting Room)";
                ws.Cells[1, 38].Value = "ARC-141 B (Small Metting Room)";
                ws.Cells[1, 39].Value = "ARC- 135 (Work Room)";
                ws.Cells[1, 40].Value = "ARC- 210 (Meeting Room)";
                ws.Cells[1, 41].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 42].Value = "Sports Field";
                ws.Cells[1, 43].Value = "Volleyball/Basketball/Tennis (combined)";
            }



            int ii = 2;
            //int rowsForAverage = 0; //Start row 2 and column 22
            //int columnsForAverage = 0;

            for (int i = 0; i < times.Length; i++)
            {
                ws.Cells[ii, 1].Value = times[i];
                ws.Cells[ii, column].Value = times[i];
                ii++;
            }

            int row = 2;
            int col = 2;
            int col_2 = 21;

            if (fileSelected == 0)
            {
                col_2 = 21;
            }
            else if (fileSelected == 1)
            {
                col_2 = 9;
            }
            else if (fileSelected == 2)
            {
                col_2 = 24;
            }

            for (int j = 0; j < roomsData.Length; j++)
            {
                for (int k = 0; k < roomsData[j].Length; k++)
                {
                    ws.Cells[row, col].Value = roomsData[j][k];
                    ws.Cells[row, col_2].Value = roomsData[j][k];

                    col++;
                    col_2++;
                }

                if (fileSelected == 0)
                {
                    col = 2;
                    col_2 = 21;
                }
                else if (fileSelected == 1)
                {
                    col = 2;
                    col_2 = 9;
                }
                else if (fileSelected == 2)
                {
                    col = 2;
                    col_2 = 24;
                }
                
                row++;
            }

            wb.Save();

            double value = 0;
            int rowsLength = 0;
            int columnStartPoint = 0;
            int columnEndPoint = 0;

            if (fileSelected == 0)
            {
                rowsLength = 17;
                columnStartPoint = 21;
                columnEndPoint = 36;
            }
            else if (fileSelected == 1)
            {
                rowsLength = 18;
                columnStartPoint = 9;
                columnEndPoint = 14;
            }
            else if (fileSelected == 2)
            {
                rowsLength = 17;
                columnStartPoint = 24;
                columnEndPoint = 44;
            }

            for (int p = 2; p <= rowsLength; p++)
            {
                for (int q = columnStartPoint; q < columnEndPoint; q++)
                {
                    value = ws.Cells[p, q].Value;
                    value = value / numberOfDays;
                    ws.Cells[p, q].Value = Math.Round(value);
                }
            }

            //To get the average just divide each cell by the number of days within the for loop, or use another for loop and create a new data field to display a seperate graph.
            //Also you may notice that the sum of the students isn't 'suming' up. That's becuase if there is an inproper entry for a time like '5:23 PM' it will be skipped, thus not adding to the totals.

            object misValue = System.Reflection.Missing.Value;
            string chart1_Range_1 = "A1";
            string chart1_Range_2 = "P17";

            string chart2_Range_1 = "T1";
            string chart2_Range_2 = "AI17";

            if (fileSelected == 0)
            {
                chart1_Range_1 = "A1";
                chart1_Range_2 = "P17";

                chart2_Range_1 = "T1";
                chart2_Range_2 = "AI17";
            }
            else if (fileSelected == 1)
            {
                chart1_Range_1 = "A1";
                chart1_Range_2 = "F18";

                chart2_Range_1 = "H1";
                chart2_Range_2 = "M18";
            }
            else if (fileSelected == 2)
            {
                chart1_Range_1 = "A1";
                chart1_Range_2 = "U17";

                chart2_Range_1 = "W1";
                chart2_Range_2 = "AQ17";
            }

            range = ws.Range[chart1_Range_1, chart1_Range_2];
            ChartObjects xlCharts = (ChartObjects)ws.ChartObjects(Type.Missing);
            //                                      Position X, Position Y, Width, Height
            ChartObject myChart = (ChartObject)xlCharts.Add(100, 350, 550, 300);
            Chart chartPage = myChart.Chart;
            chartPage.SetSourceData(range);
            chartPage.ChartType = XlChartType.xlColumnStacked;
            chartPage.ChartWizard(Source: range, Title: "Sum of all Students", CategoryTitle: "Time of Day", ValueTitle: "Number of Students");

            range = ws.Range[chart2_Range_1, chart2_Range_2];
            ChartObjects xlCharts_1 = (ChartObjects)ws.ChartObjects(Type.Missing);
            //                                      Position X, Position Y, Width, Height
            ChartObject myChart_1 = (ChartObject)xlCharts.Add(700, 350, 550, 300);
            Chart chartPage_1 = myChart_1.Chart;
            chartPage_1.SetSourceData(range);
            chartPage_1.ChartType = XlChartType.xlColumnStacked;
            chartPage_1.ChartWizard(Source: range, Title: "Average Students in Given Area", CategoryTitle: "Time of Day", ValueTitle: "Number of Students");

            wb.Save();
        }
        public _Excel()
        {

        }
        public _Excel(string path, int Sheet)
        {
            this.filePath = path;
            wb = xlApp.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            SetCurrentSheet(Sheet);
        }
        public _Excel(string path, int Sheet, string s)
        {
            this.filePath = path;
            wb = xlApp.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
            SetCurrentSheet(Sheet);
        }
        public void CleanUp()
        {
            wb.Close();
            xlApp.Quit();
        }
        public void NameWorkSheet(int i, string s, int fileSelected)
        {
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Worksheet)xlApp.Worksheets[i];
            ws = wb.Worksheets[i];
            worksheet.Name = s;

            if (fileSelected == 0)
            {
                ws.Cells[1, 1].Value = "Timestamp";
                ws.Cells[1, 2].Value = "Date";
                ws.Cells[1, 3].Value = "Time";
                ws.Cells[1, 4].Value = "First Floor Lobby";
                ws.Cells[1, 5].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 6].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 7].Value = "ARC 120";
                ws.Cells[1, 8].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 9].Value = "ARC 130 (Student Leader Office)";
                ws.Cells[1, 10].Value = "ARC - 131 A (Small Meeting Room)";
                ws.Cells[1, 11].Value = "ARC - 131 B (Small Meeting Room)";
                ws.Cells[1, 12].Value = "ARC - 141 A (Small Meeting Room)";
                ws.Cells[1, 13].Value = "ARC - 141 B (Small Meeting Room)";
                ws.Cells[1, 14].Value = "ARC - 135 (Work Room)";
                ws.Cells[1, 15].Value = "ARC - 210 (Meeting Room )";
                ws.Cells[1, 16].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 17].Value = "Sports Field";
                ws.Cells[1, 18].Value = "Volleyball/Basketball/Tennis (combined)";
            }
            else if (fileSelected == 1)
            {
                ws.Cells[1, 1].Value = "Timestamp";
                ws.Cells[1, 2].Value = "Date";
                ws.Cells[1, 3].Value = "Time";
                ws.Cells[1, 4].Value = "Lower Level Lobby";
                ws.Cells[1, 5].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 6].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 7].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 8].Value = " Fitness Studio";
            }
            else if (fileSelected == 2)
            {
                ws.Cells[1, 1].Value = "Timestamp";
                ws.Cells[1, 2].Value = "Date";
                ws.Cells[1, 3].Value = "Time";
                ws.Cells[1, 4].Value = "Lower Level Lobby";
                ws.Cells[1, 5].Value = "Zone 1 (Cardio Equipment)";
                ws.Cells[1, 6].Value = "Zone 2 (Keiser Strength Equipment) ";
                ws.Cells[1, 7].Value = "Zone 3 & 4 (strength equipment and free weight area)";
                ws.Cells[1, 8].Value = " Fitness Studio";
                ws.Cells[1, 9].Value = "First Floor Lobby";
                ws.Cells[1, 10].Value = "First Floor Gaming Alcove";
                ws.Cells[1, 11].Value = "ARC 110 (Meeting Room)";
                ws.Cells[1, 12].Value = "ARC 120";
                ws.Cells[1, 13].Value = "ARC 121 (The Board Room)";
                ws.Cells[1, 14].Value = "ARC 130 (Student Leader Office)";
                ws.Cells[1, 15].Value = "ARC - 131 A (Small Meeting Room)";
                ws.Cells[1, 16].Value = "ARC - 131 B (Small Meeting Room)";
                ws.Cells[1, 17].Value = "ARC - 141 A (Small Meeting Room)";
                ws.Cells[1, 18].Value = "ARC - 141 B (Small Meeting Room)";
                ws.Cells[1, 19].Value = "ARC - 135 (Work Room)";
                ws.Cells[1, 20].Value = "ARC - 210 (Meeting Room )";
                ws.Cells[1, 21].Value = "2nd Floor Multipurpose Space";
                ws.Cells[1, 22].Value = "Sports Field";
                ws.Cells[1, 23].Value = "Volleyball/Basketball/Tennis (combined)";
            }
        }
        public void SetCurrentSheet(int s)
        {
            currentSheet = s;
        }
        public int GetCurrentSheet()
        {
            return currentSheet;
        }
        public _Excel(string path, int sheet, int overload)
        {
            int i = 1;

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel.Sheets worksheets = xlWorkbook.Worksheets;

            var xlNewSheet = (Excel.Worksheet)worksheets.Add();
            xlNewSheet.Name = i.ToString();
            xlNewSheet.Cells[1, 1] = "New sheet content";

            xlNewSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(sheet);
            xlNewSheet.Select();

            xlWorkbook.Save();
            xlWorkbook.Close();

            releaseObject(xlNewSheet);
            releaseObject(worksheets);
            releaseObject(xlWorkbook);
            releaseObject(xlApp);

            MessageBox.Show("New Worksheet Created!");
            i++;
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public void CreateNewSheet()
        {
            Worksheet tempsheet = wb.Worksheets.Add(After: ws);
        }
        public void CreateNewFile()
        {
            this.wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            this.ws = wb.Worksheets[1];
        }
        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2;
            }
            else
            {
                return "";
            }
        }
        public void FormatCells()
        {
            range = ws.Columns["A:A"];
            range.NumberFormat = "yyyy/MM/dd";
            range = ws.Columns["B:B"];
            range.NumberFormat = "MM-dd-yyyy";
            range = ws.Columns["C:C"];
            range.NumberFormat = "hh:mm";
        }
        public void FormatColumn_A()
        {
            range = ws.Columns["A:A"];
            range.NumberFormat = "f";
        }
        public void CreateAndNameNewSheet()//int sheet
        {
            var newSheet = wb.Worksheets.Add(Type.Missing, wb.Worksheets[wb.Worksheets.Count], 1, XlSheetType.xlWorksheet) as Worksheet;
            newSheet.Name = "myWorkSheet";
        }
        //(Starting Row, Starting Column, Ending Row, Ending Column) 
        public object[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            object[,] holder = range.Value;
            object[,] returnstring = new object[endi - starti + 1, endy - starty + 1];  

            for (int p = 1; p <= endi - starti + 1; p++)
            {
                for (int q = 1; q <= endy - starty + 1; q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q];
                }
            }

            return returnstring;
        }
        public void WriteRange(int starti, int starty, int endi, int endy, object[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }
        public void WriteToCell(int i, int j, string s)
        {
            i++;
            j++;
            ws.Cells[i, j].Value2 = s;
        }
        public void FormatDates(int startR, int endR)
        {
            try
            {
                string s = "";
                string ss = "";
                Form1 form1 = new Form1();

                for (int i = startR; i <= endR; i++)
                {
                    s = ws.Cells[i, 1].Value;
                    int nameLength = s.Length;

                    if (nameLength == 25)
                    {
                        ss = s.Remove(10, 15);
                    }
                    else if (nameLength == 26)
                    {
                        ss = s.Remove(10, 16);
                    }

                    ws.Cells[i, 1].Value = ss;
                }

                wb.Save();
                wb.Close();
            }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }
            }

        }
        public void FormatDates_2(int startR, int endR)
        {
            try
            {
                string s = "";
                string ss = "";
                Form1 form1 = new Form1();

                for (int i = startR; i <= endR; i++)
                {
                    s = ws.Cells[i, 1].Value;
                    int nameLength = s.Length;

                    if (nameLength == 25)
                    {
                        ss = s.Remove(10, 15);
                    }
                    else if (nameLength == 26)
                    {
                        ss = s.Remove(10, 16);
                    }

                    ws.Cells[i, 1].Value = ss;
                }

                wb.Save();
            }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return;
                }
            }

        }
        public void MeetingRooms(int row, int bottom_W)
        {
            int col = 10;
            string s = "";

            for (int i = row; i <= bottom_W; i++)
            {
                for (int j = col; j < 14; j++)
                {
                    s = ws.Cells[i, j].Value;

                    if (s == "Un-Occupied")
                    {
                        ws.Cells[i, j].Value = 0;
                    }
                    else if (s == "Occupied")
                    {
                        ws.Cells[i, j].Value = 1;
                    }

                    col++;
                }
                col = 10;
            }

            wb.Save();
        }
        public void MeetingRooms_2(int bottom_W, int fileSelected)
        {
            int startCol = 0;
            int endCol = 0;
            string s = "";

            if (fileSelected == 0)
            {
                startCol = 10;
                endCol = 14;
            }
            else if (fileSelected == 2)
            {
                startCol = 15;
                endCol = 19;
            }

            for (int i = 2; i <= bottom_W; i++)
            {
                for (int j = startCol; j < endCol; j++)
                {
                    s = ws.Cells[i, j].Value;

                    if (s == "Un-Occupied")
                    {
                        ws.Cells[i, j].Value = 0;
                    }
                    else if (s == "Occupied")
                    {
                        ws.Cells[i, j].Value = 1;
                    }

                    startCol++;
                }

                if (fileSelected == 0)
                {
                    startCol = 10;
                }
                else if (fileSelected == 2)
                {
                    startCol = 15;
                }
            }

            wb.Save();
        }
        public double SumOfRows(int row, int bottom_W, int endColumn, int fileSelected)
        {
            double value = 0;
            int col = 4;
            int endingColumn = endColumn;
            endingColumn++;
            double sum = 0;
            double grandTotal = 0;

            try
            {
                for (int i = row; i <= bottom_W; i++)
                {
                    for (; col < endingColumn; col++)
                    {
                        value = ws.Cells[i, col].Value;
                        sum += value;
                    }

                    if (fileSelected == 0)
                    {
                        ws.Cells[i, 20].Value = "Sum =";
                        ws.Cells[i, 21].Value = sum;
                    }
                    else if (fileSelected == 1)
                    {
                        ws.Cells[i, 10].Value = "Sum =";
                        ws.Cells[i, 11].Value = sum;
                    }
                    else if (fileSelected == 2)
                    {
                        ws.Cells[i, 26].Value = "Sum =";
                        ws.Cells[i, 27].Value = sum;
                    }

                    grandTotal += sum;
                    sum = 0;
                    col = 4;
                }

                if (fileSelected == 0)
                {
                    ws.Cells[21, 20].Value = "Grand Total =";
                    ws.Cells[21, 21].Value = grandTotal;
                }
                else if (fileSelected == 1)
                {
                    ws.Cells[21, 10].Value = "Grand Total =";
                    ws.Cells[21, 11].Value = grandTotal;
                }
                else if (fileSelected == 2)
                {
                    ws.Cells[21, 26].Value = "Grand Total =";
                    ws.Cells[21, 27].Value = grandTotal;
                }
                

                wb.Save();
                return grandTotal;
                //wb.Close();
            }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    col++;
                }

                return grandTotal;
            }
        }
        public double[] ProccessArrays(double[] times)
        {
            double[] foundTimes = new double[36];

            double toCompare = times[0];
            foundTimes[0] = toCompare;
            bool duplicateWasFound = false;
            int k = 1;

            for ( int i = 0; i < times.Length; i++ )
            {
                for (int j = 0; j < foundTimes.Length; j++)
                {
                    toCompare = times[i];

                    duplicateWasFound = FindAnyDuplicates(foundTimes, toCompare);

                    if (duplicateWasFound == false && k < foundTimes.Length)
                    {
                        foundTimes[k] = toCompare;
                        k++;
                    }
                }
            }

            double[] d = foundTimes.Distinct().ToArray();
            Array.Sort(foundTimes);
            return foundTimes;
        }
        public bool FindAnyDuplicates(double[] foundTimes, double findThis)
        {
            for (int i = 0; i < foundTimes.Length; i++)
            {
                if (foundTimes[i] == findThis)
                {
                    return true;
                }
            }

            return false;
        }
        public double[][] Algorithm_JaggedArray(double[] times, int startRow, int endRow, int fileSelected)
        {
            int arraySize = 0;
            int numberOfRooms = 0;
            int endingColumn = 0;

            switch (fileSelected)
            {
                case 0:
                    arraySize = 16;
                    numberOfRooms = 15;
                    endingColumn = 19;
                    break;
                case 1:
                    arraySize = 17;
                    numberOfRooms = 5;
                    endingColumn = 9;
                    break;
                case 2:
                    arraySize = 16;
                    numberOfRooms = 20;
                    endingColumn = 24;
                    break;
            }

            double[][] arr = new double[arraySize][];

            for (int k = 0; k < arr.Length; k++)
            {
                arr[k] = new double[numberOfRooms];
            }

            double value = 0;
            int col = 4;
            double time = 0;
            int foundIndexForTime = 0;
            int indexForColumns = 0;
            int timeIndex = 0;
            int i = startRow;

            try
            {
                for (; i <= endRow; i++)
                {
                    time = ws.Cells[i, 3].Value;
                    foundIndexForTime = FindTimeofDay(time, fileSelected);
                    if (foundIndexForTime == int.MaxValue)
                    {
                        if (fileSelected == 0)
                        {
                            col = 19;
                        }
                        else if (fileSelected == 1)
                        {
                            col = 9;
                        }
                        else if (fileSelected == 2)
                        {
                            col = 24;
                        }
                        
                    }
                    else
                    {
                        col = 4;

                        switch (foundIndexForTime)
                        {
                            case 0:
                                timeIndex = 0;
                                break;
                            case 1:
                                timeIndex = 1;
                                break;
                            case 2:
                                timeIndex = 2;
                                break;
                            case 3:
                                timeIndex = 3;
                                break;
                            case 4:
                                timeIndex = 4;
                                break;
                            case 5:
                                timeIndex = 5;
                                break;
                            case 6:
                                timeIndex = 6;
                                break;
                            case 7:
                                timeIndex = 7;
                                break;
                            case 8:
                                timeIndex = 8;
                                break;
                            case 9:
                                timeIndex = 9;
                                break;
                            case 10:
                                timeIndex = 10;
                                break;
                            case 11:
                                timeIndex = 11;
                                break;
                            case 12:
                                timeIndex = 12;
                                break;
                            case 13:
                                timeIndex = 13;
                                break;
                            case 14:
                                timeIndex = 14;
                                break;
                            case 15:
                                timeIndex = 15;
                                break;
                            case 16:
                                timeIndex = 16;
                                break;
                        }
                    }

                    for (; col < endingColumn; col++)
                    {
                        value = ws.Cells[i, col].Value;
                        arr[timeIndex][indexForColumns] += value;
                        indexForColumns++;
                    }
                    indexForColumns = 0;
                    col = 4;
                }

                return arr;
            }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    //col++;

                    //Maybe build a funtion that checks the file for the user and spits back all the cells that need attention.
                    //Also don't forget to make sure that the times are being found and assigned correctly

                    MessageBox.Show("There appears to be an invalid entry in cell " + i.ToString() + ", " + col.ToString() + "\n" + "The program is halting this process until the cell has been fixed", "Invalid Cell Entry Encountered");
                    return null;
                }
                return arr;
            }
        }
        public void SetJaggedArray(double[][] thisArray)
        {
            jaggedArray = thisArray;
        }
        public double[][] GetJaggedArray()
        {
            return jaggedArray;
        }
        public int FindTimeofDay(double findThis, int fileSelected)
        {
            double[] firstFloor = new double[16];
            firstFloor[0] = 0.35416666666666669; firstFloor[1] = 0.39583333333333331; firstFloor[2] = 0.4375; firstFloor[3] = 0.47916666666666669; firstFloor[4] = 0.52083333333333337;
            firstFloor[5] = 0.5625; firstFloor[6] = 0.10416666666666667; firstFloor[7] = 0.14583333333333334; firstFloor[8] = 0.6875; firstFloor[9] = 0.72916666666666663; firstFloor[10] = 0.77083333333333337;
            firstFloor[11] = 0.8125; firstFloor[12] = 0.85416666666666663; firstFloor[13] = 0.89583333333333337; firstFloor[14] = 0.9375; firstFloor[15] = 0.97916666666666667;
            double[] otherTimes_FirstFloor = new double[16];
            otherTimes_FirstFloor[5] = 0.0625;
            otherTimes_FirstFloor[6] = 0.60416666666666663;
            otherTimes_FirstFloor[7] = 0.64583333333333337;

            double[] fitnessCenter = new double[17];
            fitnessCenter[0] = 0.3125; fitnessCenter[1] = 0.35416666666666669; fitnessCenter[2] = 0.39583333333333331; fitnessCenter[3] = 0.4375; fitnessCenter[4] = 0.47916666666666669; fitnessCenter[5] = 0.52083333333333337;
            fitnessCenter[6] = 0.5625; fitnessCenter[7] = 0.10416666666666667; fitnessCenter[8] = 0.14583333333333334; fitnessCenter[9] = 0.6875; fitnessCenter[10] = 0.72916666666666663; fitnessCenter[11] = 0.77083333333333337;
            fitnessCenter[12] = 0.8125; fitnessCenter[13] = 0.85416666666666663; fitnessCenter[14] = 0.89583333333333337; fitnessCenter[15] = 0.9375; fitnessCenter[16] = 0.97916666666666667;
            double[] otherTimes_FitnessCenter = new double[17];
            otherTimes_FitnessCenter[6] = 0.0625;
            otherTimes_FitnessCenter[7] = 0.60416666666666663;
            otherTimes_FitnessCenter[8] = 0.64583333333333337;
            otherTimes_FitnessCenter[16] = 0.97916666666666663;

            double[] infoDeskAndLowerLevel = new double[16];
            infoDeskAndLowerLevel[0] = 0.35416666666666669; infoDeskAndLowerLevel[1] = 0.39583333333333331; infoDeskAndLowerLevel[2] = 0.4375; infoDeskAndLowerLevel[3] = 0.47916666666666669; infoDeskAndLowerLevel[4] = 0.52083333333333337;
            infoDeskAndLowerLevel[5] = 0.5625; infoDeskAndLowerLevel[6] = 0.10416666666666667; infoDeskAndLowerLevel[7] = 0.14583333333333334; infoDeskAndLowerLevel[8] = 0.6875; infoDeskAndLowerLevel[9] = 0.72916666666666663; infoDeskAndLowerLevel[10] = 0.77083333333333337;
            infoDeskAndLowerLevel[11] = 0.8125; infoDeskAndLowerLevel[12] = 0.85416666666666663; infoDeskAndLowerLevel[13] = 0.89583333333333337; infoDeskAndLowerLevel[14] = 0.9375; infoDeskAndLowerLevel[15] = 0.97916666666666667;
            double[] otherTimes_InfoDeskAndLowerLevel = new double[16];
            otherTimes_InfoDeskAndLowerLevel[5] = 0.0625;
            otherTimes_InfoDeskAndLowerLevel[6] = 0.60416666666666663;
            otherTimes_InfoDeskAndLowerLevel[7] = 0.64583333333333337;

            switch (fileSelected)
            {
                case 0:

                    for (int i = 0; i < firstFloor.Length; i++)
                    {
                        if (findThis == firstFloor[i])
                        {
                            return i;
                        }
                        else if (findThis == firstFloor[i] || findThis == otherTimes_FirstFloor[i])
                        {
                            return i;
                        }
                    }

                    break;
                case 1:

                    for (int i = 0; i < fitnessCenter.Length; i++)
                    {
                        if (findThis == fitnessCenter[i])
                        {
                            return i;
                        }
                        else if (findThis == fitnessCenter[i] || findThis == otherTimes_FitnessCenter[i])
                        {
                            return i;
                        }
                    }

                    break;
                case 2:

                    for (int i = 0; i < infoDeskAndLowerLevel.Length; i++)
                    {
                        if (findThis == infoDeskAndLowerLevel[i])
                        {
                            return i;
                        }
                        else if (findThis == infoDeskAndLowerLevel[i] || findThis == otherTimes_InfoDeskAndLowerLevel[i])
                        {
                            return i;
                        }
                    }

                    break;
            }

            return int.MaxValue;
        }
        public double[] GetTimes(int top, int bottom_w, int size)
        {
            double[] time = new double[size];
            double temp = 0;
            int index = 0;
            

            for (int i = top; i <= bottom_w; i++)
            {
                temp = ws.Cells[i, 3].Value;
                time[index] = temp;
                index++;

                //Maybe try moving this whole step further up. Like before the main loop or at the start button click
            }
            return time;
        }
        public double GetThisCellValue(int top, int bottom)
        {
            double d = ws.Cells[top, bottom].Value;
            return d;
        }
        public double[] SumAllRooms(int col, int bottow_W)
        {
            double[] rooms = new double[16];
            int index = 0;
            double value = 0;
            int row = 2;

            try
            {
                for (int column = col; column < 19; column++)
                {
                    for ( ; row <= bottow_W; row++)
                    {
                        value = ws.Cells[row, column].Value;
                        rooms[index] += value;
                    }
                    index++;
                    row = 2;
                }
                return rooms;
            }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    row++;
                }
                return rooms;
            }
        }
        public void WriteRoomsToCells(double[] roomTotals, double royalTotal, int i)
        {
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Worksheet)xlApp.Worksheets[i];
            ws = wb.Worksheets[i];
            worksheet.Name = "Complete Totals";

            ws.Cells[1, 1].Value = "First Floor";
            ws.Cells[1, 2].Value = "First Floor Gaming Alcove";
            ws.Cells[1, 3].Value = "ARC 110 (Meeting Room)";
            ws.Cells[1, 4].Value = "ARC 120";
            ws.Cells[1, 5].Value = "ARC 121 (The Board Room)";
            ws.Cells[1, 6].Value = "ARC 130 (Student Leaders Office)";
            ws.Cells[1, 7].Value = "ARC-131 A (Small Metting Room)";
            ws.Cells[1, 8].Value = "ARC-131 B (Small Meeting Room)";
            ws.Cells[1, 9].Value = "ARC-141 A (Small Meeting Room)";
            ws.Cells[1, 10].Value = "ARC-141 B (Small Metting Room)";
            ws.Cells[1, 11].Value = "ARC- 135 (Work Room)";
            ws.Cells[1, 12].Value = "ARC- 210 (Meeting Room)";
            ws.Cells[1, 13].Value = "2nd Floor Multipurpose Space";
            ws.Cells[1, 14].Value = "Sports Field";
            ws.Cells[1, 15].Value = "Volleyball/Basketball/Tennis (combined)";
            ws.Cells[1, 16].Value = "Grand Total";
            ws.Cells[1, 16].Font.Size = 20;

            ws.Cells[2, 1].Value = roomTotals[0];
            ws.Cells[2, 2].Value = roomTotals[1];
            ws.Cells[2, 3].Value = roomTotals[2];
            ws.Cells[2, 4].Value = roomTotals[3];
            ws.Cells[2, 5].Value = roomTotals[4];
            ws.Cells[2, 6].Value = roomTotals[5];
            ws.Cells[2, 7].Value = roomTotals[6];
            ws.Cells[2, 8].Value = roomTotals[7];
            ws.Cells[2, 9].Value = roomTotals[8];
            ws.Cells[2, 10].Value = roomTotals[9];
            ws.Cells[2, 11].Value = roomTotals[10];
            ws.Cells[2, 12].Value = roomTotals[11];
            ws.Cells[2, 13].Value = roomTotals[12];
            ws.Cells[2, 14].Value = roomTotals[13];
            ws.Cells[2, 15].Value = roomTotals[14];
            ws.Cells[2, 16].Value = royalTotal;
            ws.Cells[2, 16].Font.Size = 20;


            range = ws.Range["A1", "O2"];
            ChartObjects xlCharts = (ChartObjects)ws.ChartObjects(Type.Missing);
            //                                      Position X, Position Y, Width, Height
            ChartObject myChart = (ChartObject)xlCharts.Add(0, 100, 1000, 600);
            Chart chartPage = myChart.Chart;
            chartPage.SetSourceData(range);
            chartPage.ChartType = XlChartType.xl3DColumnClustered;
            chartPage.ChartWizard(Source: range, Title: "Total Occupants", CategoryTitle: "Room", ValueTitle: "Number of occupants");
            //chartPage.Export(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Chart.PNG", "PNG", misValue);

            ChartObject myChart_1 = (ChartObject)xlCharts.Add(1000, 100, 1000, 600);
            Chart chartPage_1 = myChart_1.Chart;
            chartPage_1.SetSourceData(range);
            chartPage_1.ChartType = XlChartType.xlBarClustered;
            chartPage_1.ChartWizard(Source: range, Title: "Total Occupants", CategoryTitle: "Room", ValueTitle: "Number of occupants");

            ChartObject myChart_2 = (ChartObject)xlCharts.Add(500, 700, 1000, 600);
            Chart chartPage_2 = myChart_2.Chart;
            chartPage_2.SetSourceData(range);
            chartPage_2.ChartType = XlChartType.xl3DPie;
            chartPage_2.ChartWizard(Source: range, Title: "Total Occupants");

            //ChartObject myChart_3 = (ChartObject)xlCharts.Add(1000, 1000, 1000, 600);
            //Chart chartPage_3 = myChart_3.Chart;
            //chartPage_3.SetSourceData(range);
            //chartPage_3.ChartType = XlChartType.xlColumnStacked;
            //chartPage_3.ChartWizard(Source: range, Title: "Total Occupants", CategoryTitle: "Room", ValueTitle: "Number of occupants");

            //ChartObject myChart_4 = (ChartObject)xlCharts.Add(0, 1600, 1000, 600);
            //Chart chartPage_4 = myChart_4.Chart;
            //chartPage_4.SetSourceData(range);
            //chartPage_4.ChartType = XlChartType.xlLineMarkers;
            //chartPage_4.ChartWizard(Source: range, Title: "Total Occupants", CategoryTitle: "Room", ValueTitle: "Number of occupants");
            wb.Save();
        }
        public void AssignRowIndentifier(int rows)
        {
            ws.Cells[22, 20].Value = "Critical Data for ARC HEAD COUNTS Software System, do not move or delete the adjacent cell";
            ws.Cells[22, 21].Value = rows;

            //You made is so only when the checkbox is not checked the identifier should be written
        }
        public string[] RunPreProcessTest(int startRow, int lastRow, int fileSelected)
        {
            double value = 0;
            int row = startRow;
            int col = 4;
            int index = 0;
            string badCellLocation = "";
            string[] cellsToFix = new string[100];
            bool run = true;

            int endingCol = 0;

            switch (fileSelected)
            {
                case 0:
                    endingCol = 19;
                    break;
                case 1:
                    endingCol = 9;
                    break;
                case 2:
                    endingCol = 24;
                    break;
            }

            while (run == true)
            {
                try
                {
                    for (; row <= lastRow; row++)
                    {
                        for (; col < endingCol; col++)
                        {
                            value = ws.Cells[row, col].Value;
                        }
                        col = 4;
                    }

                    run = false;
                }
                catch (Exception exception)
                {
                    if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        badCellLocation = row.ToString() + ", " + col.ToString();
                        cellsToFix[index] = badCellLocation;
                        index++;

                        int tempRow = row;
                        int tempCol = col;

                        if (col == 18 & tempRow++ != lastRow)
                        {
                            row = tempRow++;
                        }
                        else
                        {
                            tempRow--;
                            row = tempRow;
                        }

                        if (tempCol != 18)
                        {
                            col = ++tempCol;
                        }
                        else
                        {
                            col = 4;
                        }
                    }
                }
            }

            return cellsToFix;
        }
        public bool CorrectFileLoaded(int fileSelected)
        {
            string value_1 = "";
            string value_2 = "";

            if (fileSelected == 0)
            {
                value_1 = ws.Cells[1, 4].Value;
                value_2 = ws.Cells[1, 18].Value;

                if (value_1 == "First Floor Lobby" && value_2 == "Volleyball/Basketball/Tennis (combined)")
                {
                    return true;
                }
            }
            else if (fileSelected == 1)
            {
                value_1 = ws.Cells[1, 4].Value;
                value_2 = ws.Cells[1, 8].Value;

                if (value_1 == "Number of participants in the Lower Level Lobby?" && value_2 == "Number of participants in Fitness Studio" && ws.Cells[1, 9].Value != "What is the Fitness Studio currently being used for?")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else if (fileSelected == 2)
            {
                value_1 = ws.Cells[1, 4].Value;
                value_2 = ws.Cells[1, 23].Value;

                if (value_1 == "Lower Level Lobby:" && value_2 == "Volleyball/Basketball/Tennis (combined)" && ws.Cells[1, 9].Value != "What is the Fitness Studio currently being used for?")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return false;
        }
        public void CreateHelperTextFile(string[] cellToFix)
        {
            StreamWriter output = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Fix_these_cells.txt");

            var time = DateTime.Now;
            string formmattedTime = time.ToString("MM/dd/yyyy - hh:mm:ss tt");
            output.WriteLine("Generated: " + formmattedTime);

            output.WriteLine();

            output.WriteLine("Row, Column");

            for (int i = 0; i < cellToFix.Length; i++)
            {
                if (cellToFix[i] != String.Empty)
                {
                    output.WriteLine(cellToFix[i]);
                }
            }

            

            output.Dispose();
        }
        public void Save()
        {
            wb.Save();
        }
        public void SaveAs(string path)
        {
            xlApp.DisplayAlerts = false;
            wb.SaveAs(path + ".xlsx");
            xlApp.DisplayAlerts = true;
        }
        public void Close()
        {
            wb.Close();
        }
        public void CLoseAndQuit()
        {
            wb.Close(0);
            xlApp.Quit();
        }
        public void SelectWorksheet(int SheetNumber)
        {
            this.ws = wb.Worksheets[SheetNumber];
        }
        public void DeleteWorksheet(int SheetNumber)
        {
            wb.Worksheets[SheetNumber].Delete();
        }
        public void ProtectSheet()
        {
            ws.Protect();
        }
        public void ProtectSheet(string Password)
        {
            ws.Protect(Password);
        }
        public void UnprotectSheet()
        {
            ws.Unprotect();
        }
        public void UnprotectSheet(string Password)
        {
            ws.Unprotect(Password);
        }
    }
}