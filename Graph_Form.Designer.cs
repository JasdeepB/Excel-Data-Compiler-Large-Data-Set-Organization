namespace ARC_Head_Counts
{
    partial class Graph_Form
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
            this.pictureBoxGraphs = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGraphs)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxGraphs
            // 
            this.pictureBoxGraphs.Location = new System.Drawing.Point(1, 1);
            this.pictureBoxGraphs.Name = "pictureBoxGraphs";
            this.pictureBoxGraphs.Size = new System.Drawing.Size(1233, 760);
            this.pictureBoxGraphs.TabIndex = 0;
            this.pictureBoxGraphs.TabStop = false;
            // 
            // Graph_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1234, 761);
            this.Controls.Add(this.pictureBoxGraphs);
            this.Name = "Graph_Form";
            this.Text = "View Graph";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxGraphs)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.PictureBox pictureBoxGraphs;
    }
}