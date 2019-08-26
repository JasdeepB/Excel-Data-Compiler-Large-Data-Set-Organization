namespace ARC_Head_Counts
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.sheetLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.startingDateTxtb = new System.Windows.Forms.TextBox();
            this.endingRangeTxtB = new System.Windows.Forms.TextBox();
            this.startingRangeBtn = new System.Windows.Forms.Button();
            this.endingDateBtn = new System.Windows.Forms.Button();
            this.testFileBtn = new System.Windows.Forms.Button();
            this.customRangeBtn = new System.Windows.Forms.Button();
            this.selectAndSaveBtn = new System.Windows.Forms.Button();
            this.newFileNameTxtBx = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.getSpecificCellValueBtn = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.columnSelection = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.testNewSheetBtn = new System.Windows.Forms.Button();
            this.formatTestBtn = new System.Windows.Forms.Button();
            this.creatorsName = new System.Windows.Forms.Label();
            this.creatorLinkedIn = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.fileTypeLabel = new System.Windows.Forms.Label();
            this.fileTypeCombobox = new System.Windows.Forms.ComboBox();
            this.createGraphBtn = new System.Windows.Forms.Button();
            this.graphType = new System.Windows.Forms.Label();
            this.checkBox = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.currentFileTxtb = new System.Windows.Forms.TextBox();
            this.clearDateBtn_2 = new System.Windows.Forms.Button();
            this.clearDateBtn_1 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.sumAndAveBtn = new System.Windows.Forms.Button();
            this.comboBoxGraph = new System.Windows.Forms.ComboBox();
            this.openBtn = new System.Windows.Forms.Button();
            this.startBtn = new System.Windows.Forms.Button();
            this.Graph = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.stopBtn = new System.Windows.Forms.Button();
            this.testCellsBtn = new System.Windows.Forms.Button();
            this.graphBtn = new System.Windows.Forms.Button();
            this.testSumOfRowsBtn = new System.Windows.Forms.Button();
            this.cellFinderBtn = new System.Windows.Forms.Button();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.currentStateMessage = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.BackgroundColor = System.Drawing.Color.Black;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dataGridView.GridColor = System.Drawing.Color.Black;
            this.dataGridView.Location = new System.Drawing.Point(319, 0);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 62;
            this.dataGridView.RowTemplate.DefaultCellStyle.NullValue = null;
            this.dataGridView.Size = new System.Drawing.Size(1031, 881);
            this.dataGridView.TabIndex = 0;
            // 
            // comboBox
            // 
            this.comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox.FormattingEnabled = true;
            this.comboBox.Location = new System.Drawing.Point(31, 349);
            this.comboBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(261, 24);
            this.comboBox.TabIndex = 1;
            this.comboBox.SelectedIndexChanged += new System.EventHandler(this.ComboBox_SelectedIndexChanged);
            // 
            // sheetLabel
            // 
            this.sheetLabel.AutoSize = true;
            this.sheetLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sheetLabel.ForeColor = System.Drawing.Color.White;
            this.sheetLabel.Location = new System.Drawing.Point(131, 327);
            this.sheetLabel.Name = "sheetLabel";
            this.sheetLabel.Size = new System.Drawing.Size(57, 20);
            this.sheetLabel.TabIndex = 2;
            this.sheetLabel.Text = "Sheet";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(96, 379);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "Set Date Range";
            // 
            // startingDateTxtb
            // 
            this.startingDateTxtb.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.startingDateTxtb.Location = new System.Drawing.Point(60, 405);
            this.startingDateTxtb.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.startingDateTxtb.Name = "startingDateTxtb";
            this.startingDateTxtb.Size = new System.Drawing.Size(189, 23);
            this.startingDateTxtb.TabIndex = 5;
            // 
            // endingRangeTxtB
            // 
            this.endingRangeTxtB.Location = new System.Drawing.Point(60, 436);
            this.endingRangeTxtB.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.endingRangeTxtB.Name = "endingRangeTxtB";
            this.endingRangeTxtB.Size = new System.Drawing.Size(188, 23);
            this.endingRangeTxtB.TabIndex = 6;
            // 
            // startingRangeBtn
            // 
            this.startingRangeBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.startingRangeBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.startingRangeBtn.ForeColor = System.Drawing.Color.White;
            this.startingRangeBtn.Location = new System.Drawing.Point(256, 405);
            this.startingRangeBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.startingRangeBtn.Name = "startingRangeBtn";
            this.startingRangeBtn.Size = new System.Drawing.Size(33, 24);
            this.startingRangeBtn.TabIndex = 7;
            this.startingRangeBtn.Text = "OK";
            this.startingRangeBtn.UseVisualStyleBackColor = true;
            this.startingRangeBtn.Click += new System.EventHandler(this.StartingRangeBtn_Click);
            // 
            // endingDateBtn
            // 
            this.endingDateBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.endingDateBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.endingDateBtn.ForeColor = System.Drawing.Color.White;
            this.endingDateBtn.Location = new System.Drawing.Point(256, 437);
            this.endingDateBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.endingDateBtn.Name = "endingDateBtn";
            this.endingDateBtn.Size = new System.Drawing.Size(33, 23);
            this.endingDateBtn.TabIndex = 8;
            this.endingDateBtn.Text = "OK";
            this.endingDateBtn.UseVisualStyleBackColor = true;
            this.endingDateBtn.Click += new System.EventHandler(this.EndingDateBtn_Click);
            // 
            // testFileBtn
            // 
            this.testFileBtn.Location = new System.Drawing.Point(420, 18);
            this.testFileBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.testFileBtn.Name = "testFileBtn";
            this.testFileBtn.Size = new System.Drawing.Size(146, 44);
            this.testFileBtn.TabIndex = 9;
            this.testFileBtn.Text = "TestFile()";
            this.testFileBtn.UseVisualStyleBackColor = true;
            this.testFileBtn.Click += new System.EventHandler(this.TestFileBtn_Click);
            // 
            // customRangeBtn
            // 
            this.customRangeBtn.Location = new System.Drawing.Point(12, 78);
            this.customRangeBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.customRangeBtn.Name = "customRangeBtn";
            this.customRangeBtn.Size = new System.Drawing.Size(122, 44);
            this.customRangeBtn.TabIndex = 11;
            this.customRangeBtn.Text = "Test Custom Range";
            this.customRangeBtn.UseVisualStyleBackColor = true;
            this.customRangeBtn.Click += new System.EventHandler(this.CustomRangeBtn_Click);
            // 
            // selectAndSaveBtn
            // 
            this.selectAndSaveBtn.Location = new System.Drawing.Point(420, 78);
            this.selectAndSaveBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.selectAndSaveBtn.Name = "selectAndSaveBtn";
            this.selectAndSaveBtn.Size = new System.Drawing.Size(146, 44);
            this.selectAndSaveBtn.TabIndex = 12;
            this.selectAndSaveBtn.Text = "SelectAndSave()";
            this.selectAndSaveBtn.UseVisualStyleBackColor = true;
            this.selectAndSaveBtn.Click += new System.EventHandler(this.SelectAndSaveBtn_Click);
            // 
            // newFileNameTxtBx
            // 
            this.newFileNameTxtBx.Location = new System.Drawing.Point(32, 494);
            this.newFileNameTxtBx.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.newFileNameTxtBx.Name = "newFileNameTxtBx";
            this.newFileNameTxtBx.Size = new System.Drawing.Size(261, 23);
            this.newFileNameTxtBx.TabIndex = 13;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(118, 467);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 20);
            this.label2.TabIndex = 14;
            this.label2.Text = "File Name";
            // 
            // getSpecificCellValueBtn
            // 
            this.getSpecificCellValueBtn.Location = new System.Drawing.Point(140, 78);
            this.getSpecificCellValueBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.getSpecificCellValueBtn.Name = "getSpecificCellValueBtn";
            this.getSpecificCellValueBtn.Size = new System.Drawing.Size(122, 44);
            this.getSpecificCellValueBtn.TabIndex = 15;
            this.getSpecificCellValueBtn.Text = "GetSpecificCellValue()";
            this.getSpecificCellValueBtn.UseVisualStyleBackColor = true;
            this.getSpecificCellValueBtn.Click += new System.EventHandler(this.GetSpecificCellValueBtn_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(319, 881);
            this.progressBar.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(706, 70);
            this.progressBar.TabIndex = 16;
            this.progressBar.Visible = false;
            // 
            // columnSelection
            // 
            this.columnSelection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.columnSelection.FormattingEnabled = true;
            this.columnSelection.Location = new System.Drawing.Point(32, 669);
            this.columnSelection.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.columnSelection.Name = "columnSelection";
            this.columnSelection.Size = new System.Drawing.Size(261, 24);
            this.columnSelection.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(134, 645);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 20);
            this.label3.TabIndex = 18;
            this.label3.Text = "Room";
            // 
            // testNewSheetBtn
            // 
            this.testNewSheetBtn.Location = new System.Drawing.Point(1102, 58);
            this.testNewSheetBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.testNewSheetBtn.Name = "testNewSheetBtn";
            this.testNewSheetBtn.Size = new System.Drawing.Size(148, 28);
            this.testNewSheetBtn.TabIndex = 20;
            this.testNewSheetBtn.Text = "Test New Sheet Name";
            this.testNewSheetBtn.UseVisualStyleBackColor = true;
            this.testNewSheetBtn.Click += new System.EventHandler(this.TestNewSheetBtn_Click);
            // 
            // formatTestBtn
            // 
            this.formatTestBtn.Location = new System.Drawing.Point(1256, 58);
            this.formatTestBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.formatTestBtn.Name = "formatTestBtn";
            this.formatTestBtn.Size = new System.Drawing.Size(148, 28);
            this.formatTestBtn.TabIndex = 21;
            this.formatTestBtn.Text = "Test Formatting";
            this.formatTestBtn.UseVisualStyleBackColor = true;
            this.formatTestBtn.Click += new System.EventHandler(this.FormatTestBtn_Click);
            // 
            // creatorsName
            // 
            this.creatorsName.AutoSize = true;
            this.creatorsName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.creatorsName.ForeColor = System.Drawing.Color.White;
            this.creatorsName.Location = new System.Drawing.Point(1031, 897);
            this.creatorsName.Name = "creatorsName";
            this.creatorsName.Size = new System.Drawing.Size(319, 20);
            this.creatorsName.TabIndex = 23;
            this.creatorsName.Text = "Created With Passion By Jasdeep Brar";
            // 
            // creatorLinkedIn
            // 
            this.creatorLinkedIn.AutoSize = true;
            this.creatorLinkedIn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.creatorLinkedIn.ForeColor = System.Drawing.Color.White;
            this.creatorLinkedIn.Location = new System.Drawing.Point(1055, 921);
            this.creatorLinkedIn.Name = "creatorLinkedIn";
            this.creatorLinkedIn.Size = new System.Drawing.Size(272, 15);
            this.creatorLinkedIn.TabIndex = 24;
            this.creatorLinkedIn.Text = "https://www.linkedin.com/in/jasdeep-brar/";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(205, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(138, 20);
            this.label4.TabIndex = 25;
            this.label4.Text = "Developer Tools";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.fileTypeLabel);
            this.panel1.Controls.Add(this.fileTypeCombobox);
            this.panel1.Controls.Add(this.createGraphBtn);
            this.panel1.Controls.Add(this.graphType);
            this.panel1.Controls.Add(this.checkBox);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.currentFileTxtb);
            this.panel1.Controls.Add(this.clearDateBtn_2);
            this.panel1.Controls.Add(this.clearDateBtn_1);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.sumAndAveBtn);
            this.panel1.Controls.Add(this.comboBoxGraph);
            this.panel1.Controls.Add(this.openBtn);
            this.panel1.Controls.Add(this.comboBox);
            this.panel1.Controls.Add(this.columnSelection);
            this.panel1.Controls.Add(this.startBtn);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.sheetLabel);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.startingDateTxtb);
            this.panel1.Controls.Add(this.newFileNameTxtBx);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.endingRangeTxtB);
            this.panel1.Controls.Add(this.startingRangeBtn);
            this.panel1.Controls.Add(this.endingDateBtn);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(320, 951);
            this.panel1.TabIndex = 27;
            // 
            // fileTypeLabel
            // 
            this.fileTypeLabel.AutoSize = true;
            this.fileTypeLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileTypeLabel.ForeColor = System.Drawing.Color.White;
            this.fileTypeLabel.Location = new System.Drawing.Point(121, 144);
            this.fileTypeLabel.Name = "fileTypeLabel";
            this.fileTypeLabel.Size = new System.Drawing.Size(81, 20);
            this.fileTypeLabel.TabIndex = 34;
            this.fileTypeLabel.Text = "File Type";
            // 
            // fileTypeCombobox
            // 
            this.fileTypeCombobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fileTypeCombobox.FormattingEnabled = true;
            this.fileTypeCombobox.Location = new System.Drawing.Point(31, 167);
            this.fileTypeCombobox.Name = "fileTypeCombobox";
            this.fileTypeCombobox.Size = new System.Drawing.Size(260, 24);
            this.fileTypeCombobox.TabIndex = 33;
            this.fileTypeCombobox.SelectedIndexChanged += new System.EventHandler(this.FileTypeCombobox_SelectedIndexChanged);
            // 
            // createGraphBtn
            // 
            this.createGraphBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(44)))), ((int)(((byte)(51)))));
            this.createGraphBtn.Cursor = System.Windows.Forms.Cursors.Default;
            this.createGraphBtn.FlatAppearance.BorderSize = 0;
            this.createGraphBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.createGraphBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.createGraphBtn.ForeColor = System.Drawing.Color.White;
            this.createGraphBtn.Image = ((System.Drawing.Image)(resources.GetObject("createGraphBtn.Image")));
            this.createGraphBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.createGraphBtn.Location = new System.Drawing.Point(0, 852);
            this.createGraphBtn.Name = "createGraphBtn";
            this.createGraphBtn.Size = new System.Drawing.Size(320, 87);
            this.createGraphBtn.TabIndex = 30;
            this.createGraphBtn.Text = "Get Graph";
            this.createGraphBtn.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.createGraphBtn.UseVisualStyleBackColor = true;
            this.createGraphBtn.Click += new System.EventHandler(this.CreateGraphBtn_Click);
            // 
            // graphType
            // 
            this.graphType.AutoSize = true;
            this.graphType.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.graphType.ForeColor = System.Drawing.Color.White;
            this.graphType.Location = new System.Drawing.Point(116, 799);
            this.graphType.Name = "graphType";
            this.graphType.Size = new System.Drawing.Size(102, 20);
            this.graphType.TabIndex = 32;
            this.graphType.Text = "Graph Type";
            // 
            // checkBox
            // 
            this.checkBox.AutoSize = true;
            this.checkBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox.ForeColor = System.Drawing.Color.White;
            this.checkBox.Location = new System.Drawing.Point(56, 524);
            this.checkBox.Name = "checkBox";
            this.checkBox.Size = new System.Drawing.Size(233, 24);
            this.checkBox.TabIndex = 31;
            this.checkBox.Text = "Show Graphs for All Days";
            this.checkBox.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(114, 280);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(103, 20);
            this.label5.TabIndex = 30;
            this.label5.Text = "Current File";
            // 
            // currentFileTxtb
            // 
            this.currentFileTxtb.Location = new System.Drawing.Point(32, 303);
            this.currentFileTxtb.Name = "currentFileTxtb";
            this.currentFileTxtb.Size = new System.Drawing.Size(260, 23);
            this.currentFileTxtb.TabIndex = 29;
            // 
            // clearDateBtn_2
            // 
            this.clearDateBtn_2.FlatAppearance.BorderSize = 0;
            this.clearDateBtn_2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.clearDateBtn_2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearDateBtn_2.ForeColor = System.Drawing.Color.White;
            this.clearDateBtn_2.Location = new System.Drawing.Point(26, 436);
            this.clearDateBtn_2.Name = "clearDateBtn_2";
            this.clearDateBtn_2.Size = new System.Drawing.Size(29, 23);
            this.clearDateBtn_2.TabIndex = 28;
            this.clearDateBtn_2.Text = "X";
            this.clearDateBtn_2.UseVisualStyleBackColor = true;
            this.clearDateBtn_2.Click += new System.EventHandler(this.ClearDateBtn_2_Click);
            // 
            // clearDateBtn_1
            // 
            this.clearDateBtn_1.FlatAppearance.BorderSize = 0;
            this.clearDateBtn_1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.clearDateBtn_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearDateBtn_1.ForeColor = System.Drawing.Color.White;
            this.clearDateBtn_1.Location = new System.Drawing.Point(26, 404);
            this.clearDateBtn_1.Name = "clearDateBtn_1";
            this.clearDateBtn_1.Size = new System.Drawing.Size(29, 23);
            this.clearDateBtn_1.TabIndex = 27;
            this.clearDateBtn_1.Text = "X";
            this.clearDateBtn_1.UseVisualStyleBackColor = true;
            this.clearDateBtn_1.Click += new System.EventHandler(this.ClearDateBtn_1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(320, 137);
            this.pictureBox1.TabIndex = 22;
            this.pictureBox1.TabStop = false;
            // 
            // sumAndAveBtn
            // 
            this.sumAndAveBtn.FlatAppearance.BorderSize = 0;
            this.sumAndAveBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.sumAndAveBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sumAndAveBtn.ForeColor = System.Drawing.Color.White;
            this.sumAndAveBtn.Image = ((System.Drawing.Image)(resources.GetObject("sumAndAveBtn.Image")));
            this.sumAndAveBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.sumAndAveBtn.Location = new System.Drawing.Point(0, 701);
            this.sumAndAveBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.sumAndAveBtn.Name = "sumAndAveBtn";
            this.sumAndAveBtn.Size = new System.Drawing.Size(320, 87);
            this.sumAndAveBtn.TabIndex = 19;
            this.sumAndAveBtn.Text = "Get Sum and Average";
            this.sumAndAveBtn.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.sumAndAveBtn.UseVisualStyleBackColor = true;
            this.sumAndAveBtn.Click += new System.EventHandler(this.SumAndAveBtn_Click);
            // 
            // comboBoxGraph
            // 
            this.comboBoxGraph.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxGraph.FormattingEnabled = true;
            this.comboBoxGraph.Location = new System.Drawing.Point(32, 822);
            this.comboBoxGraph.Name = "comboBoxGraph";
            this.comboBoxGraph.Size = new System.Drawing.Size(261, 24);
            this.comboBoxGraph.TabIndex = 29;
            // 
            // openBtn
            // 
            this.openBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(44)))), ((int)(((byte)(51)))));
            this.openBtn.FlatAppearance.BorderSize = 0;
            this.openBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.openBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openBtn.ForeColor = System.Drawing.Color.White;
            this.openBtn.Image = ((System.Drawing.Image)(resources.GetObject("openBtn.Image")));
            this.openBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.openBtn.Location = new System.Drawing.Point(0, 191);
            this.openBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.openBtn.Name = "openBtn";
            this.openBtn.Size = new System.Drawing.Size(320, 85);
            this.openBtn.TabIndex = 3;
            this.openBtn.Text = "Open";
            this.openBtn.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.openBtn.UseVisualStyleBackColor = false;
            this.openBtn.Click += new System.EventHandler(this.OpenBtn_Click);
            // 
            // startBtn
            // 
            this.startBtn.FlatAppearance.BorderSize = 0;
            this.startBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.startBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.startBtn.ForeColor = System.Drawing.Color.White;
            this.startBtn.Image = ((System.Drawing.Image)(resources.GetObject("startBtn.Image")));
            this.startBtn.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.startBtn.Location = new System.Drawing.Point(2, 549);
            this.startBtn.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.startBtn.Name = "startBtn";
            this.startBtn.Size = new System.Drawing.Size(317, 92);
            this.startBtn.TabIndex = 10;
            this.startBtn.Text = "Start";
            this.startBtn.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.startBtn.UseVisualStyleBackColor = true;
            this.startBtn.Click += new System.EventHandler(this.StartBtn_Click);
            // 
            // Graph
            // 
            this.Graph.Location = new System.Drawing.Point(140, 18);
            this.Graph.Name = "Graph";
            this.Graph.Size = new System.Drawing.Size(122, 44);
            this.Graph.TabIndex = 33;
            this.Graph.Text = "Test Making Graph";
            this.Graph.UseVisualStyleBackColor = true;
            this.Graph.Click += new System.EventHandler(this.Graph_Click);
            // 
            // panel2
            // 
            this.panel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel2.Controls.Add(this.stopBtn);
            this.panel2.Controls.Add(this.Graph);
            this.panel2.Controls.Add(this.testCellsBtn);
            this.panel2.Controls.Add(this.graphBtn);
            this.panel2.Controls.Add(this.testSumOfRowsBtn);
            this.panel2.Controls.Add(this.cellFinderBtn);
            this.panel2.Controls.Add(this.label4);
            this.panel2.Controls.Add(this.formatTestBtn);
            this.panel2.Controls.Add(this.testFileBtn);
            this.panel2.Controls.Add(this.testNewSheetBtn);
            this.panel2.Controls.Add(this.getSpecificCellValueBtn);
            this.panel2.Controls.Add(this.customRangeBtn);
            this.panel2.Controls.Add(this.selectAndSaveBtn);
            this.panel2.Location = new System.Drawing.Point(519, 444);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(580, 153);
            this.panel2.TabIndex = 28;
            this.panel2.Visible = false;
            // 
            // stopBtn
            // 
            this.stopBtn.Location = new System.Drawing.Point(453, 125);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(75, 23);
            this.stopBtn.TabIndex = 33;
            this.stopBtn.Text = "Stop";
            this.stopBtn.UseVisualStyleBackColor = true;
            this.stopBtn.Click += new System.EventHandler(this.StopBtn_Click);
            // 
            // testCellsBtn
            // 
            this.testCellsBtn.Location = new System.Drawing.Point(12, 18);
            this.testCellsBtn.Name = "testCellsBtn";
            this.testCellsBtn.Size = new System.Drawing.Size(122, 44);
            this.testCellsBtn.TabIndex = 30;
            this.testCellsBtn.Text = "Test Getting Cells Value";
            this.testCellsBtn.UseVisualStyleBackColor = true;
            this.testCellsBtn.Click += new System.EventHandler(this.TestCellsBtn_Click);
            // 
            // graphBtn
            // 
            this.graphBtn.Location = new System.Drawing.Point(1102, 13);
            this.graphBtn.Name = "graphBtn";
            this.graphBtn.Size = new System.Drawing.Size(148, 28);
            this.graphBtn.TabIndex = 28;
            this.graphBtn.Text = "Graph()";
            this.graphBtn.UseVisualStyleBackColor = true;
            this.graphBtn.Click += new System.EventHandler(this.GraphBtn_Click);
            // 
            // testSumOfRowsBtn
            // 
            this.testSumOfRowsBtn.Location = new System.Drawing.Point(268, 78);
            this.testSumOfRowsBtn.Name = "testSumOfRowsBtn";
            this.testSumOfRowsBtn.Size = new System.Drawing.Size(146, 44);
            this.testSumOfRowsBtn.TabIndex = 27;
            this.testSumOfRowsBtn.Text = "TestSumOfRows()";
            this.testSumOfRowsBtn.UseVisualStyleBackColor = true;
            this.testSumOfRowsBtn.Click += new System.EventHandler(this.TestSumOfRowsBtn_Click);
            // 
            // cellFinderBtn
            // 
            this.cellFinderBtn.Location = new System.Drawing.Point(268, 18);
            this.cellFinderBtn.Name = "cellFinderBtn";
            this.cellFinderBtn.Size = new System.Drawing.Size(146, 44);
            this.cellFinderBtn.TabIndex = 26;
            this.cellFinderBtn.Text = "TestCellFinder()";
            this.cellFinderBtn.UseVisualStyleBackColor = true;
            this.cellFinderBtn.Click += new System.EventHandler(this.CellFinderBtn_Click);
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(319, 881);
            this.progressBar2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(706, 70);
            this.progressBar2.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar2.TabIndex = 29;
            this.progressBar2.Visible = false;
            // 
            // currentStateMessage
            // 
            this.currentStateMessage.AutoSize = true;
            this.currentStateMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentStateMessage.ForeColor = System.Drawing.Color.White;
            this.currentStateMessage.Location = new System.Drawing.Point(1031, 885);
            this.currentStateMessage.Name = "currentStateMessage";
            this.currentStateMessage.Size = new System.Drawing.Size(178, 20);
            this.currentStateMessage.TabIndex = 31;
            this.currentStateMessage.Text = "ARC HEAD COUNTS";
            this.currentStateMessage.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(41)))), ((int)(((byte)(44)))), ((int)(((byte)(51)))));
            this.ClientSize = new System.Drawing.Size(1350, 951);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.currentStateMessage);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.creatorLinkedIn);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.creatorsName);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Font = new System.Drawing.Font("Bahnschrift", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximumSize = new System.Drawing.Size(1366, 990);
            this.MinimumSize = new System.Drawing.Size(638, 354);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ARC Head Counts";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.Label sheetLabel;
        private System.Windows.Forms.Button openBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox startingDateTxtb;
        private System.Windows.Forms.TextBox endingRangeTxtB;
        private System.Windows.Forms.Button startingRangeBtn;
        private System.Windows.Forms.Button endingDateBtn;
        private System.Windows.Forms.Button testFileBtn;
        private System.Windows.Forms.Button startBtn;
        private System.Windows.Forms.Button customRangeBtn;
        private System.Windows.Forms.Button selectAndSaveBtn;
        private System.Windows.Forms.TextBox newFileNameTxtBx;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button getSpecificCellValueBtn;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.ComboBox columnSelection;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button sumAndAveBtn;
        private System.Windows.Forms.Button testNewSheetBtn;
        private System.Windows.Forms.Button formatTestBtn;
        private System.Windows.Forms.Label creatorsName;
        private System.Windows.Forms.Label creatorLinkedIn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button clearDateBtn_1;
        private System.Windows.Forms.Button clearDateBtn_2;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox currentFileTxtb;
        private System.Windows.Forms.Button cellFinderBtn;
        private System.Windows.Forms.Button testSumOfRowsBtn;
        private System.Windows.Forms.Button graphBtn;
        private System.Windows.Forms.Button createGraphBtn;
        private System.Windows.Forms.ComboBox comboBoxGraph;
        public System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.CheckBox checkBox;
        private System.Windows.Forms.Label graphType;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.Button testCellsBtn;
        private System.Windows.Forms.Label currentStateMessage;
        private System.Windows.Forms.Button Graph;
        private System.Windows.Forms.Button stopBtn;
        private System.Windows.Forms.Label fileTypeLabel;
        public System.Windows.Forms.ComboBox fileTypeCombobox;
    }
}