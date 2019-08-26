using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using BackedExcelFunctions;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Diagnostics;

namespace ARC_Head_Counts
{
    public partial class Form1 : Form
    {
        string filePath = @"";
        string newFileName = @"";
        string dateReference = "";
        string fileToDelete = "";
        string tempFilePath = "";
        string[] times;
        DataSet result;
        Range range;
        int starti;
        int starty;
        int endi;
        int endy;
        int worksheetIndex;
        double sum;
        double average;
        bool startDateSet = false;
        bool endDateSet = false;
        bool stopButtonClicked = false;

        _Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb;
        Worksheet ws;

        _Excel excel;

        //TO DO - Add a feature that shows what cells need to be fixed in order for the system to run through without problems - *Done*
        //Add a feature that allows the user to stop
        //Add a feature that allows the user to clear the current file
        //Add a feature that prompts the user the confirm whether or not to close the program - *Done*

        //Known Bugs
        //1. - 'Get Graph' is no longer displaying graphs - *Fixed*
        //2. - There are unique cases where files are in use, or are not closed, causing IO and COM exceptions
        //3. - Excel sheets are still open in the task manager even after being told to close, or perhaps those exact one's are yet to be told to closed yet, or at all.

        public Form1()
        {
            InitializeComponent();
            LoadGraphSelection();
            LoadFileNameSelection();
            progressBar2.Visible = false;
            newFileNameTxtBx.Text = "Debug Mode";
            //ModifyProgressBarColor.SetState(progressBar, 3);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        public void TestFile()
        {
            //(Starting Row, Starting Column, Ending Row, Ending Column)
            _Excel ex = new _Excel(@"C:\Users\uhits\OneDrive\Desktop\2018- 2019 ARC Head Counts- Information Desk Rearranged", 1);
            object[,] read = ex.ReadRange(2, 1, 4000, 18);
            ex.Close();
            _Excel ex1 = new _Excel(@"C:\Users\uhits\OneDrive\Desktop\Book1", 1);
            ex1.WriteRange(2, 1, 4000, 18, read);
            ex1.FormatCells();
            ex1.Save();
            ex1.Close();
        }
        private void OpenBtn_Click(object sender, EventArgs e)
        {
            OpenFile();
        }
        public void OpenFile()
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook| *.xlsx", ValidateNames = true })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        filePath = ofd.FileName;
                        currentFileTxtb.Text = filePath;
                        FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                        IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                        result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                        comboBox.Items.Clear();
                        foreach (System.Data.DataTable dt in result.Tables)
                        {
                            comboBox.Items.Add(dt.TableName);
                            reader.Close();
                        }
                    }
                };

                //creatorsName.Visible = false;
                //creatorLinkedIn.Visible = false;
                //currentStateMessage.Visible = true;
                //currentStateMessage.Text = "Getting Things Ready";
            }
            catch (Exception exception)
            {
                if (exception is System.IO.IOException)
                {
                    string title = "File in use";
                    string message = "That Excel file is currently in use." + "\n\n" + "If it's currently open, please close it and try again."
                    + "\n\n" + "If you don't have it open then your computer has it open in the background. End the task through the Task Manager and try again."
                    + " \n" + "Look for 'Microsoft Excel'.";
                    MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView.DataSource = result.Tables[comboBox.SelectedIndex];
            dataGridView.Columns[0].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm";
            dataGridView.Columns[1].DefaultCellStyle.Format = "MM/dd/yyyy";
            dataGridView.Columns[2].DefaultCellStyle.Format = "HH:mm";
        }
        private void StartingRangeBtn_Click(object sender, EventArgs e)
        {
            try
            {
                SetStartI(dataGridView.CurrentRow.Index);
                int columnSelected = dataGridView.CurrentCell.ColumnIndex;

                if (columnSelected > 1)
                {
                    MessageBox.Show("That is not a valid date, valid dates are considered anything from the first or second column", "Invalid date selected");
                }
                else
                {
                    if (columnSelected == 1)
                    {
                        SetStartY(0);
                        startingDateTxtb.Text = dataGridView.SelectedCells[0].Value.ToString();
                        Set_StartDate(true);
                    }
                    else
                    {
                        SetStartY(0);
                        startingDateTxtb.Text = dataGridView.SelectedCells[0].Value.ToString();
                        Set_StartDate(true);
                    }

                }
            }
            catch (System.NullReferenceException)
            {

            }
        }
        public void SetStartI(int i) //Starting Row
        {
            starti = i + 2; //If you change 'UseHeaderRow' to false change this back to + 1
        }
        public void SetStartY(int y) // Starting Column
        {
            starty = y + 1;
        }
        public void SetEndI(int i) //Ending Row
        {
            endi = i + 2; //If you change 'UseHeaderRow' to false change this back to + 1
        }
        public void SetEndY(int y) //Ending Column
        {
            endy = y + 1;
        }
        public int GetStartI()
        {
            return starti;
        }
        public int GetStartY()
        {
            return starty;
        }
        public int GetEndI()
        {
            return endi;
        }
        public int GetEndY()
        {
            return endy;
        }
        public void SetWorksheetIndex(int i)
        {
            worksheetIndex = i;
        }
        public int GetWorksheetIndex()
        {
            return worksheetIndex;
        }
        public void SetNewFileName(string fileName)
        {
            newFileName = fileName;
        }
        public string GetNewFileName()
        {
            return newFileName;
        }
        public void Set_StartDate(bool b)
        {
            startDateSet = b;
        }
        public bool Get_StartDate()
        {
            return startDateSet;
        }
        public void Set_EndDate(bool b)
        {
            endDateSet = b;
        }
        public bool Get_EndDate()
        {
            return endDateSet;
        }
        private void EndingDateBtn_Click(object sender, EventArgs e)
        {
            try
            {
                SetEndI(dataGridView.CurrentRow.Index);
                int columnSelected = dataGridView.CurrentCell.ColumnIndex;

                if (columnSelected > 1)
                {
                    MessageBox.Show("That is not a valid date, valid dates are considered anything from the first or second column", "Invalid date selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    if (columnSelected == 1)
                    {
                        SetEndY(17);
                        endingRangeTxtB.Text = dataGridView.SelectedCells[0].Value.ToString();
                        Set_EndDate(true);
                    }
                    else
                    {
                        SetEndY(17);
                        endingRangeTxtB.Text = dataGridView.SelectedCells[0].Value.ToString();
                        Set_EndDate(true);
                    }
                }
            }
            catch (System.NullReferenceException)
            {
                
            }
        }
        private async void TestFileBtn_Click(object sender, EventArgs e)
        {
            TestFile();
        }
        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (fileTypeCombobox.SelectedIndex < 0)
            {
                MessageBox.Show("The type of file has not been specified."+ "\n" +"Please select one and try agian.", "No File Type Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (checkBox.Checked == true)
            {
                string title = "Attention!";
                string message = "You've selected to have every day to contain graphs." + "\n" + "Doing so will require more processing time, about 3X more time.";
                DialogResult result = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (result == DialogResult.No)
                {
                    return;
                }
            }

            if (progressBar.Value > 0)
            {
                progressBar.Minimum = 0;
            }

            bool start = Get_StartDate();
            bool end = Get_EndDate();

            if (start == false || end == false)
            {
                if (end == false && start == true)
                {
                    MessageBox.Show("The ending date has not yet been set.", "Date Not Set", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else if (start == false && end == true)
                {
                    MessageBox.Show("The starting date has not yet been set.", "Date Not Set", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

            }

            if (start == false && end == false)
            {
                MessageBox.Show("No dates have been selected", "No Dates Given", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                if (newFileNameTxtBx.Text == String.Empty)
                {
                    MessageBox.Show("No name was given to the file, give it a name and try again", "No File Name", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    _Excel excel = new _Excel(filePath, 1);
                    int checkForValidDateEntry = GetEndI() - GetStartI();

                    if (excel.CorrectFileLoaded(fileTypeCombobox.SelectedIndex) == false)
                    {
                        string title = "Improper File Selected";
                        string message = "It appears that the wrong file type may have been selected." + "\n\n"+ "Or"+ "\n\n" + "If you believe this to be an error make sure you've set-up the file properly and removed the correct columns prior to loading it into this system.";
                        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        excel.CLoseAndQuit();

                        startingDateTxtb.Text = "";
                        Set_StartDate(false);
                        endingRangeTxtB.Text = "";
                        Set_EndDate(false);

                        return;
                    }

                    if (checkForValidDateEntry < 1)
                    {
                        string title = "Invalid Date Selection Order";
                        string message = "The order in which you've selected your dates are invalid" + "\n" + "Please try again";
                        MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        startingDateTxtb.Text = "";
                        Set_StartDate(false);
                        endingRangeTxtB.Text = "";
                        Set_EndDate(false);

                        return;
                    }
                    else
                    {
                        FileInfo file = new FileInfo(filePath);
                        string f = (Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\ARC HEAD COUNTS - Copy of Selected File ~ DO NOT DELETE.xlsx");
                        file.CopyTo(f, true);
                        tempFilePath = filePath;
                        filePath = f;
                        fileToDelete = f;

                        RunAlgorithm();
                    }

                    //RunAlgorithm();
                }
            }
        }
        public async void RunAlgorithm()
        {
            excel = new _Excel(filePath, 1);
            int arraySize = GetEndI() - GetStartI();
            int fileSelected = fileTypeCombobox.SelectedIndex;
            
            if (creatorsName.Visible == true && creatorLinkedIn.Visible == true)
            {
                creatorsName.Visible = false;
                creatorLinkedIn.Visible = false;
            }

            if (currentStateMessage.Visible == false)
            {
                currentStateMessage.Visible = true;
                currentStateMessage.Text = "Setting up file for processing" + "\n" + "Please wait..." + "\n" + "Step 1 of 2";
            }
            else
            {
                currentStateMessage.Visible = true;
                currentStateMessage.Text = "Setting up file for processing" + "\n" + "Please wait..." + "\n" + "Step 1 of 2";
            }

            if (progressBar2.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    progressBar2.Visible = true;
                });
            }
            else
            {
                progressBar2.Visible = true;
            }

            if (fileSelected == 0 || fileSelected == 2)
            {
                await Task.Factory.StartNew(() => excel.MeetingRooms_2(GetEndI(), fileSelected));
            }

            string[] cellsToFix = await Task.Factory.StartNew(() => excel.RunPreProcessTest(GetStartI(), GetEndI(), fileSelected));

            if (cellsToFix[0] != null)
            {
                await Task.Factory.StartNew(() => excel.CreateHelperTextFile(cellsToFix));
                currentStateMessage.Visible = false;
                creatorsName.Visible = true;
                creatorLinkedIn.Visible = true;
                progressBar.Visible = false;
                progressBar2.Visible = false;

                string message = "There are cells that have invalid entries." + "\n" + "They must be fixed before attempting to process again" + "\n" + "Do you want to see which need to be fixed?";
                string title = "Invalid Entries Have Been Detected";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    Process.Start("notepad.exe", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Fix_these_cells.txt");
                    return;
                }
                else
                {
                    return;
                }
            }

            double[] times = await Task.Factory.StartNew(() => excel.GetTimes(GetStartI(), GetEndI(), ++arraySize));
            await Task.Factory.StartNew(() => excel.FormatDates_2(GetStartI(), GetEndI()));
            double[] foundTimes = await Task.Factory.StartNew(() => excel.ProccessArrays(times));
            double[][] processedArray = await Task.Factory.StartNew(() => excel.Algorithm_JaggedArray(foundTimes, GetStartI(), GetEndI(), fileSelected));
            
            //Make equivilanties for standard time and military time, so that the value for 1:30 is seen as the same as 11:30 and vice versa
            if (progressBar2.InvokeRequired)
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    progressBar2.Visible = false;
                });
            }
            else
            {
                progressBar2.Visible = false;
            }

            if (processedArray != null)
            {
                currentStateMessage.Text = "Processing the file" + "\n" + "Please wait..." + "\n" + "Step 2 of 2";
                await Task.Factory.StartNew(() => Algorithm(processedArray));
            }
            else
            {
                return;
            }

            //using (var cancellationTokenSource = new CancellationTokenSource())
            //{
            //    var stopTask = Task.Run(() =>
            //    {
            //        stopBtn.Click += new System.EventHandler(StopBtn_Click);
            //    });

            //    try
            //    {

            //    }
            //    catch (TaskCanceledException)
            //    {
            //        MessageBox.Show("Returned");
            //        return;
            //    }
            //}
        }
        public void TestIndexes()
        {
            MessageBox.Show("Start Row: " + GetStartI().ToString() + "\n" + "Start Column: " + GetStartY().ToString() + "\n" + "End Row: " + GetEndI().ToString() + "\n" + "End Column: " + GetEndY().ToString());
        }
        private void CustomRangeBtn_Click(object sender, EventArgs e)
        {
            //(Starting Row, Starting Column, Ending Row, Ending Column) //starti, starty, endi, endy //GetStartI(),GetStartY(),GetEndI(),GetEndY()
            _Excel ex = new _Excel(filePath, 1);
            object[,] read = ex.ReadRange(GetStartI(), GetStartY(), GetEndI(), GetEndY());
            ex.Close();
            _Excel ex1 = new _Excel(@"C:\Users\uhits\OneDrive\Desktop\Book1", 1);
            ex1.WriteRange(GetStartI(), GetStartY(), GetEndI(), GetEndY(), read);
            ex1.FormatCells();
            ex1.Save();
            ex1.Close();

            TestIndexes();
        }
        public void SelectAndSave()
        {
            _Excel ex = new _Excel(filePath, 1);
            object[,] read = ex.ReadRange(GetStartI(), GetStartY(), GetEndI(), GetEndY());
            ex.Close();

            _Excel ex1 = new _Excel();
            ex1.CreateNewFile();
            ex1.SaveAs(newFileNameTxtBx.Text);
            SetNewFileName(newFileNameTxtBx.Text);

            string fileName = GetNewFileName() + ".xlsx";
            FileInfo f = new FileInfo(fileName);
            string fullName = f.Name;

            ex1.Close();

            _Excel ex2 = new _Excel(fullName, 1);
            ex2.WriteRange(GetStartI(), GetStartY(), GetEndI(), GetEndY(), read);
            ex2.FormatCells();
            ex2.Save();
            ex2.Close();
        }
        private void SelectAndSaveBtn_Click(object sender, EventArgs e)
        {
            SelectAndSave();
            TestIndexes();
            //TestFilePathName();
        }
        public void TestFilePathName()
        {
            _Excel ex1 = new _Excel();
            ex1.CreateNewFile();
            ex1.CreateNewSheet();
            ex1.SaveAs(newFileNameTxtBx.Text);
            SetNewFileName(newFileNameTxtBx.Text);
            ex1.Close();

            string fileName = GetNewFileName() + ".xlsx";
            FileInfo f = new FileInfo(fileName);
            string fullName = f.Name;
            MessageBox.Show(fullName);
        }
        public string GetSpecificCellValueR(int rowp, int colomnp)
        {
            try
            {
                int row = rowp;
                int column = colomnp;

                if (rowp == 1)
                {
                    row++;
                }

                string test = ws.Cells[row, column].Value.ToString();
                DateTime oDate = DateTime.Parse(test);
                string dateToPass = oDate.Month + "-" + oDate.Day + "-" + oDate.Year;
                //MessageBox.Show("Here's the date: " + GetDateReference() + "\n" + "Column: " + column + "\n" + "Row: " + row);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return dateToPass;
            }
            catch(Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    return "";
                }
                else
                {
                    return "";
                }
            }
        }
        public void GetSpecificCellValue(int rowp, int colomnp)
        {
            _Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];

            int row = rowp;
            int column = colomnp;

            if (rowp == 1)
            {
                row++;
            }

            string test = ws.Cells[row, column].Value.ToString();
            DateTime oDate = DateTime.Parse(test);
            MessageBox.Show(oDate.Day + "-" + oDate.Month + "-" + oDate.Year);
            //SetDateReference(test);
            //MessageBox.Show("Here's the date: " + GetDateReference() + "\n" + "Column: " + column + "\n" + "Row: " + row);
            xlApp.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void SetDateReference(string date)
        {
            dateReference = date;
        }
        public string GetDateReference()
        {
            return dateReference;
        }
        private void GetSpecificCellValueBtn_Click(object sender, EventArgs e)
        {
            //for (int i = 1; i < 22; i++)
            //{
            //    string t = GetSpecificCellValueR(i, 2);
            //    MessageBox.Show(t);
            //}
            GetSpecificCellValue(GetStartI(), 2);
        }
        public void StringComparisonTest()
        {
            _Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(filePath);
            Worksheet ws = wb.Worksheets[1];
            string test1 = ws.Cells[2, 2].Value.ToString();
            string test2 = ws.Cells[2356, 2].Value.ToString();
            
            if (test1 == test2)
            {
                MessageBox.Show("The dates are the same" + "\n" + "First date: " + test1 + "\n" + "Second Date: " + test2);
            }
            else
            {
                MessageBox.Show("Yes the two date are different" + "\n" + "First date: " + test1 + "\n" + "Second Date: " + test2);
            }
            xlApp.Quit();
        }
        public void LoadFirstFloorRooms()
        {
            columnSelection.Items.Add("First Floor Lobby"); //0
            columnSelection.Items.Add("First Floor Gaming Alcove"); //1
            columnSelection.Items.Add("ARC 110 (Meeting Room)"); //2
            columnSelection.Items.Add("ARC 120"); //3
            columnSelection.Items.Add("ARC 121 (The Board Room)"); //4
            columnSelection.Items.Add("ARC 130 (Student Leader Office)"); //5
            columnSelection.Items.Add("ARC- 135 (Work Room)"); //6
            columnSelection.Items.Add("ARC- 210 (Meeting Room )"); //7
            columnSelection.Items.Add("2nd Floor Multipurpose Space"); //8
            columnSelection.Items.Add("Sports Field"); //9
            columnSelection.Items.Add("Volleyball/Basketball/Tennis (combined)"); //10
        }
        public void LoadLowerLevelRooms()
        {
            columnSelection.Items.Add("Lower Level Lobby");
            columnSelection.Items.Add("Zone 1 (Cardio Equipment)");
            columnSelection.Items.Add("Zone 2 (Keiser Strength Equipment)");
            columnSelection.Items.Add("Zone 3 & 4 (strength equipment and free weight area)");
            columnSelection.Items.Add("Fitness Studio");
        }
        public void LoadInfoDeskAndLowerLevelRooms()
        {
            columnSelection.Items.Add("Lower Level Lobby");
            columnSelection.Items.Add("Zone 1 (Cardio Equipment)");
            columnSelection.Items.Add("Zone 2 (Keiser Strength Equipment)");
            columnSelection.Items.Add("Zone 3 & 4 (strength equipment and free weight area)");
            columnSelection.Items.Add("Fitness Studio");
            columnSelection.Items.Add("First Floor Lobby"); 
            columnSelection.Items.Add("First Floor Gaming Alcove"); 
            columnSelection.Items.Add("ARC 110 (Meeting Room)"); 
            columnSelection.Items.Add("ARC 120"); 
            columnSelection.Items.Add("ARC 121 (The Board Room)"); 
            columnSelection.Items.Add("ARC 130 (Student Leader Office)");
            columnSelection.Items.Add("ARC- 135 (Work Room)");
            columnSelection.Items.Add("ARC- 210 (Meeting Room )");
            columnSelection.Items.Add("2nd Floor Multipurpose Space");
            columnSelection.Items.Add("Sports Field");
            columnSelection.Items.Add("Volleyball/Basketball/Tennis (combined)");
        }
        public void LoadFileNameSelection()
        {
            fileTypeCombobox.Items.Add("1 - Information Desk Head Count");
            fileTypeCombobox.Items.Add("2 - Fitness Center Head Count");
            fileTypeCombobox.Items.Add("3 - Standard Head Count");
        }
        public void LoadGraphSelection()
        {
            comboBoxGraph.Items.Add("3D Column");
            comboBoxGraph.Items.Add("Area Column");
            comboBoxGraph.Items.Add("Column Clustered");
            comboBoxGraph.Items.Add("Column Stacked");
            comboBoxGraph.Items.Add("Line Markers");
        }
        public void GetSumAndAverage()
        {
            try
            {
                xlApp.Visible = false;
                wb = xlApp.Workbooks.Open(filePath);
                int index = comboBox.SelectedIndex;
                ws = wb.Worksheets[++index];
                double count = dataGridView.Rows.Count;

                if (fileTypeCombobox.SelectedIndex == 0)
                {
                    switch (columnSelection.SelectedIndex)
                    {
                        case 0:
                            range = ws.Range["D:D"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "First Floor Looby");
                            wb.Close();
                            break;
                        case 1:
                            range = ws.Range["E:E"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "First Floor Gaming Alcove");
                            break;
                        case 2:
                            range = ws.Range["F:F"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 110 (Meeting Room)");
                            break;
                        case 3:
                            range = ws.Range["G:G"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 120");
                            break;
                        case 4:
                            range = ws.Range["H:H"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 121 (The Board Room)");
                            break;
                        case 5:
                            range = ws.Range["I:I"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 130 (Student Leader Office)");
                            break;
                        case 6:
                            range = ws.Range["N:N"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC- 135 (Work Room)");
                            break;
                        case 7:
                            range = ws.Range["O:O"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC- 210 (Meeting Room )");
                            break;
                        case 8:
                            range = ws.Range["P:P"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "2nd Floor Multipurpose Space");
                            break;
                        case 9:
                            range = ws.Range["Q:Q"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Sports Field");
                            break;
                        case 10:
                            range = ws.Range["R:R"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Volleyball/Basketball/Tennis (combined)");
                            break;
                    }
                }
                else if (fileTypeCombobox.SelectedIndex == 1)
                {
                    switch (columnSelection.SelectedIndex)
                    {
                        case 0:
                            range = ws.Range["D:D"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Lower Level Lobby");
                            wb.Close();
                            break;
                        case 1:
                            range = ws.Range["E:E"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 1 (Cardio Equipment)");
                            wb.Close();
                            break;
                        case 2:
                            range = ws.Range["F:F"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 2 (Keiser Strength Equipment)");
                            wb.Close();
                            break;
                        case 3:
                            range = ws.Range["G:G"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 3 & 4 (strength equipment and free weight area)");
                            wb.Close();
                            break;
                        case 4:
                            range = ws.Range["H:H"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Fitness Studio");
                            wb.Close();
                            break;
                    }
                }
                else if (fileTypeCombobox.SelectedIndex == 2)
                {
                    switch (columnSelection.SelectedIndex)
                    {
                        case 0:
                            range = ws.Range["D:D"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Lower Level Lobby");
                            wb.Close();
                            break;
                        case 1:
                            range = ws.Range["E:E"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 1 (Cardio Equipment)");
                            wb.Close();
                            break;
                        case 2:
                            range = ws.Range["F:F"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 2 (Keiser Strength Equipment)");
                            wb.Close();
                            break;
                        case 3:
                            range = ws.Range["G:G"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Zone 3 & 4 (strength equipment and free weight area)");
                            wb.Close();
                            break;
                        case 4:
                            range = ws.Range["H:H"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Fitness Studio");
                            wb.Close();
                            break;
                        case 5:
                            range = ws.Range["I:I"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "First Floor Looby");
                            wb.Close();
                            break;
                        case 6:
                            range = ws.Range["J:J"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "First Floor Gaming Alcove");
                            break;
                        case 7:
                            range = ws.Range["K:K"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 110 (Meeting Room)");
                            break;
                        case 8:
                            range = ws.Range["L:L"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 120");
                            break;
                        case 9:
                            range = ws.Range["M:M"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 121 (The Board Room)");
                            break;
                        case 10:
                            range = ws.Range["N:N"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC 130 (Student Leader Office)");
                            break;
                        case 11:
                            range = ws.Range["S:S"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC- 135 (Work Room)");
                            break;
                        case 12:
                            range = ws.Range["T:T"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "ARC- 210 (Meeting Room )");
                            break;
                        case 13:
                            range = ws.Range["U:U"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "2nd Floor Multipurpose Space");
                            break;
                        case 14:
                            range = ws.Range["V:V"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Sports Field");
                            break;
                        case 15:
                            range = ws.Range["W:W"];
                            sum = xlApp.WorksheetFunction.Sum(range);
                            average = xlApp.WorksheetFunction.Average(range);
                            MessageBox.Show("The sum is " + sum.ToString() + "\n" + "The average is " + average.ToString(), "Volleyball/Basketball/Tennis (combined)");
                            break;
                    }
                }
            }
            catch(Exception exception)
            {
                if (exception is System.Runtime.InteropServices.COMException)
                {
                    string message = "No file is currently selected" + "\n" + "Would you like to open one now?";
                    string title = "No file selected";
                    var result = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        OpenFile();
                    }
                    else if (result == DialogResult.No)
                    {
                        return;
                    }
                }
            }          
        }
        public void Algorithm(double[][] roomsData)
        {
            try
            {
                wb = xlApp.Workbooks.Open(filePath);
                ws = wb.Worksheets[1];

                int topRange = GetStartI();
                int bottomRange = topRange;
                int bottomRange_W = 2;
                int numberofDays = 0;
                int columnsWide = 18;
                int fileSelected = 0;

                if (fileTypeCombobox.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        if (fileTypeCombobox.SelectedIndex == 1)
                        {
                            columnsWide = 8;
                            fileSelected = 1;
                        }
                        else if (fileTypeCombobox.SelectedIndex == 2)
                        {
                            columnsWide = 23;
                            fileSelected = 2;
                        }
                    });
                }

                string referenceString = GetSpecificCellValueR(topRange, 1);
                string stringForSheet = "";
                int sheet = 1;
                double[] roomTotals = new double[16];

                if (fileTypeCombobox.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                       if (fileTypeCombobox.SelectedIndex == 0 || fileTypeCombobox.SelectedIndex == 2)
                       {
                            times = new string[16];

                            times[0] = "8:30 AM"; times[1] = "9:30 AM"; times[2] = "10:30 AM"; times[3] = "11:30 AM"; times[4] = "12:30 PM"; times[5] = "1:30 PM"; times[6] = "2:30 PM"; times[7] = "3:30 PM";
                            times[8] = "4:30 PM"; times[9] = "5:30 PM"; times[10] = "6:30 PM"; times[11] = "7:30 PM"; times[12] = "8:30 PM"; times[13] = "9:30 PM"; times[14] = "10:30 PM"; times[15] = "11:30 PM";
                       }
                       else if (fileTypeCombobox.SelectedIndex == 1)
                       {
                            times = new string[17];

                            times[0] = "7:30 AM"; times[1] = "8:30 AM"; times[2] = "9:30 AM"; times[3] = "10:30 AM"; times[4] = "11:30 AM"; times[5] = "12:30 PM"; times[6] = "1:30 PM"; times[7] = "2:30 PM";
                            times[8] = "3:30 PM"; times[9] = "4:30 PM"; times[10] = "5:30 PM"; times[11] = "6:30 PM"; times[12] = "7:30 PM"; times[13] = "8:30 PM"; times[14] = "9:30 PM"; times[15] = "10:30 PM"; times[16] = "11:30 PM";
                       }
                    });
                }

                

                _Excel ex1 = new _Excel();
                ex1.CreateNewFile();
                ex1.SaveAs(newFileNameTxtBx.Text);
                SetNewFileName(newFileNameTxtBx.Text);
                ex1.CLoseAndQuit();

                _Excel ex = new _Excel(filePath, 1);
                object[,] read = ex.ReadRange(GetStartI(), 1, GetEndI(), 18);

                string fileName = GetNewFileName() + ".xlsx";
                FileInfo f = new FileInfo(fileName);
                string fullName = f.Name;

                _Excel ex2 = new _Excel();
                ex2 = new _Excel(fullName, sheet);

                if (progressBar.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        if (progressBar.Value > 0)
                        {
                            progressBar.Value = 0;
                        }

                        progressBar.Minimum = GetStartI();
                        progressBar.Maximum = GetEndI();
                        progressBar.Step = 1;
                        progressBar.Visible = true;
                    });
                }

                for (int i = GetStartI(); i <= GetEndI(); i++)
                {
                    int num = i;

                    if (num != GetStartI())
                    {
                        referenceString = GetSpecificCellValueR(++num, 1);
                        num--;
                        stringForSheet = GetSpecificCellValueR(num, 1);
                    }
                    else
                    {
                        referenceString = GetSpecificCellValueR(GetStartI(), 1);
                    }

                    if (GetSpecificCellValueR(i, 1) == referenceString)
                    {
                        bottomRange = i;
                        bottomRange_W++;
                    }
                    else
                    {
                        if (checkBox.Checked == true)
                        {
                            bottomRange = i;

                            //If you make changes to ReadRange or WriteRange you must insure that both have the same numbers  
                            read = ex.ReadRange(topRange, 1, bottomRange, columnsWide);

                            ex2.NameWorkSheet(sheet, stringForSheet, fileSelected);
                            ex2.WriteRange(2, 1, bottomRange_W, columnsWide, read);//////DO NOT DELETE
                            ex2.SumOfRows(2, bottomRange_W, columnsWide, fileSelected);
                            ex2.FormatCells();
                            if (fileTypeCombobox.InvokeRequired)
                            {
                                this.Invoke((MethodInvoker)delegate ()
                                {
                                    ex2.Graph(bottomRange_W, fileTypeCombobox.SelectedIndex);
                                });
                            }
                            ex2.CreateNewSheet();/////DO NOT DELETE
                            ex2.Save();

                            topRange = bottomRange;
                            topRange++;
                            sheet++;
                            bottomRange_W = 2;
                            numberofDays++;

                        }
                        else
                        {
                            bottomRange = i;

                            //If you make changes to ReadRange or WriteRange you must insure that both have the same numbers  
                            read = ex.ReadRange(topRange, 1, bottomRange, columnsWide);

                            ex2.NameWorkSheet(sheet, stringForSheet, fileSelected);
                            ex2.WriteRange(2, 1, bottomRange_W, columnsWide, read);//////DO NOT DELETE
                            ex2.AssignRowIndentifier(bottomRange_W);
                            ex2.SumOfRows(2, bottomRange_W, columnsWide, fileSelected);
                            ex2.FormatCells();
                            ex2.CreateNewSheet();/////DO NOT DELETE
                            ex2.Save();

                            topRange = bottomRange;
                            topRange++;
                            sheet++;
                            bottomRange_W = 2;
                            numberofDays++;
                        }
                    }


                    if (progressBar.InvokeRequired)
                    {
                        this.Invoke((MethodInvoker)delegate ()
                        {
                            progressBar.PerformStep();
                        });
                    }
                }

                if (currentStateMessage.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        currentStateMessage.Text = "Finsihing up";
                    });
                }

                ex2.GenerateSummaryGraph(times, roomsData, numberofDays, sheet, fileSelected);

                ex.CLoseAndQuit();

                ex2.CLoseAndQuit();

                wb.Close();
                xlApp.Quit();

                if (progressBar.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        progressBar.Increment(100);
                    });
                }

                excel.CLoseAndQuit();
                
  
                FileInfo file = new FileInfo(fileToDelete);
                file.Delete();
                filePath = tempFilePath;

                string title = "All Done!";
                string message = "Your new Workbook named " + "'" + GetNewFileName() + "'" + " was saved in your Documents folder";
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (currentStateMessage.InvokeRequired && creatorsName.InvokeRequired && creatorLinkedIn.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        currentStateMessage.Visible = false;
                        creatorsName.Visible = true;
                        creatorLinkedIn.Visible = true;

                    });
                }

                if (progressBar.InvokeRequired)
                {
                    this.Invoke((MethodInvoker)delegate ()
                    {
                        progressBar.Visible = false;
                    });
                }
        }
            catch (Exception exception)
            {
                if (exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    string title = "Previously compiled book";
                    string message = "It appears the date range contains dates that have been previously processed by this system." + "\n" + "You can avoid this error by picking a date range that hasn't been processed already.";
                    MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                else
                {
                    MessageBox.Show(exception.ToString(), "Exception was Thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void SumAndAveBtn_Click(object sender, EventArgs e)
        {
            GetSumAndAverage();
        }
        public void ShowNewWorksheetName(string s)
        {
            MessageBox.Show(s);
        }
        private void TestNewSheetBtn_Click(object sender, EventArgs e)
        {
            //CreateNewSheet_Named();
            wb = xlApp.Workbooks.Open(@"C:\Users\uhits\Documents\Test Book.xlsx");
            
            int j = 1;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Worksheet)xlApp.Worksheets[j];
            ws = wb.Worksheets[j];

            for (int i = 1; i < 5; i++)
            {
                worksheet = wb.Worksheets.Add(Before: ws);
                //Did not create a new sheet, make a new sheet.
                worksheet.Name = i.ToString() + "-" + i.ToString() + "-" + "2019";
                MessageBox.Show(worksheet.Name);
                
                j++;
            }
            wb.Save();
            wb.Close();   
        }
        public void TestFormatting()
        {
            _Excel excel = new _Excel(filePath, 1);
            excel.FormatDates(GetStartI(), GetEndI());
        }
        private void FormatTestBtn_Click(object sender, EventArgs e)
        {
            TestFormatting();
            //TestStringLength();
            //dataGridView.Columns[1].DefaultCellStyle.Format = "MM-dd-yyyy";
        }
        public void FormatDataGridView()
        {
            dataGridView.Columns[1].DefaultCellStyle.Format = "MM-dd-yyyy";
        }
        private void ClearDateBtn_1_Click(object sender, EventArgs e)
        {
            startingDateTxtb.Text = "";
            SetStartI(0);
            SetStartY(0);
            Set_StartDate(false);
        }
        private void ClearDateBtn_2_Click(object sender, EventArgs e)
        {
            endingRangeTxtB.Text = "";
            SetEndI(0);
            SetEndY(0);
            Set_EndDate(false);
        }
        public void TestStringLength()
        {
            wb = xlApp.Workbooks.Open(@"C:\Users\uhits\Desktop\2018- 2019 ARC Head Counts- Information Desk.xlsx");
            ws = wb.Worksheets[1];

            string s = ws.Cells[3, 1].Value;
            int nameLength = s.Length;
            MessageBox.Show(s+ "\n" + nameLength.ToString());
        }
        public void TestCellFinder(int row, int column)
        {
            wb = xlApp.Workbooks.Open(@"C:\Users\uhits\Desktop\2018- 2019 ARC Head Counts- Information Desk.xlsx");
            ws = wb.Worksheets[1];

            string s = ws.Cells[row, column].Value;
            MessageBox.Show(s);
        }
        private void CellFinderBtn_Click(object sender, EventArgs e)
        {
            //TestCellFinder(2, 10);
            _Excel ex = new _Excel(filePath, 1);
            ex.MeetingRooms(2, 3170);
        }
        private void TestSumOfRowsBtn_Click(object sender, EventArgs e)
        {
            _Excel excel = new _Excel(filePath, 1);
            excel.SumOfRows(2, 14, 8, 0);
        }
        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            xlApp.Quit();
        }
        private void GraphBtn_Click(object sender, EventArgs e)
        {
            _Excel excel = new _Excel(filePath, 1);
            //excel.Graph();
        }
        private void CreateGraphBtn_Click(object sender, EventArgs e)
        {
            int indexForGraph = comboBoxGraph.SelectedIndex;
            int indexForWorksheet = comboBox.SelectedIndex;
            indexForWorksheet++;
            indexForGraph++;
            _Excel excel = new _Excel(filePath, indexForWorksheet);
            excel.ShowGraph(indexForGraph, indexForWorksheet, fileTypeCombobox.SelectedIndex);
            excel.Close();
        }
        private void TestCellsBtn_Click(object sender, EventArgs e)
        {
            _Excel excel = new _Excel(@"C:\\Users\\uhits\\Desktop\\Test\\Time Tester\\Tester.xlsx", 1);
            //double[] time = excel.TestCellValue(15, 30);
            //var time = excel.GetThisCellValue(708, 3);
            
            //MessageBox.Show(TimeSpan.FromMinutes(time).ToString());
            MessageBox.Show(excel.GetThisCellValue(938,3).ToString());

        }
        private void Graph_Click(object sender, EventArgs e)
        {
            _Excel excel = new _Excel(filePath, 1);
            excel.TestGraph();
        }
        private void StopBtn_Click(object sender, EventArgs e)
        {
            Set_StopBtnClicked(true);
        }
        public void Set_StopBtnClicked(bool clicked)
        {
            stopButtonClicked = clicked;
        }
        public bool Get_StopBtnClicked()
        {
            return stopButtonClicked;
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            DialogResult result = MessageBox.Show("Do you really want to exit the program?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            try
            {
                if (result == DialogResult.Yes)
                {
                    xlApp.Quit();
                }
                else if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            catch (System.InvalidCastException)
            {

            }
        }
        private void FileTypeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fileTypeCombobox.SelectedIndex == 0)
            {
                LoadFirstFloorRooms();
                columnSelection.Items.Clear();
                LoadFirstFloorRooms();
            }
            else if (fileTypeCombobox.SelectedIndex == 1)
            {
                LoadLowerLevelRooms();
                columnSelection.Items.Clear();
                LoadLowerLevelRooms();
            }
            else if (fileTypeCombobox.SelectedIndex == 2)
            {
                LoadInfoDeskAndLowerLevelRooms();
                columnSelection.Items.Clear();
                LoadInfoDeskAndLowerLevelRooms();
            }
        }
    }
}