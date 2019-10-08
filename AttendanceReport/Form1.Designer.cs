namespace AttendanceReport
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
            this.tbxPNumber.KeyPress -= TextBox1_KeyPress;
            this.tbxCarNumber.KeyPress -= TextBox1_KeyPress;

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
            this.btnGenerate = new System.Windows.Forms.Button();
            this.dtpFromDate = new System.Windows.Forms.DateTimePicker();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dtpToDate = new System.Windows.Forms.DateTimePicker();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dtpLateTimeStart = new System.Windows.Forms.DateTimePicker();
            this.gbxLateTimeRange = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpLateTimeEnd = new System.Windows.Forms.DateTimePicker();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.tbxCnic = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cbxCompany = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cbxCadre = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cbxCrew = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tbxCarNumber = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbxPNumber = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbxName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cbxSections = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbxDepartments = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.gbxLateTimeRange.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(593, 319);
            this.btnGenerate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(151, 59);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "Generate Report";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.button1_Click);
            // 
            // dtpFromDate
            // 
            this.dtpFromDate.CustomFormat = "";
            this.dtpFromDate.Location = new System.Drawing.Point(8, 21);
            this.dtpFromDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpFromDate.Name = "dtpFromDate";
            this.dtpFromDate.Size = new System.Drawing.Size(260, 22);
            this.dtpFromDate.TabIndex = 1;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "pdf";
            this.saveFileDialog1.Filter = "PDF|*.pdf|Excel|*.xlsx";
            this.saveFileDialog1.Title = "Select Path To Save Report";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtpFromDate);
            this.groupBox1.Location = new System.Drawing.Point(19, 255);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(280, 53);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "From Date";
            // 
            // dtpToDate
            // 
            this.dtpToDate.CustomFormat = "";
            this.dtpToDate.Location = new System.Drawing.Point(8, 21);
            this.dtpToDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpToDate.Name = "dtpToDate";
            this.dtpToDate.Size = new System.Drawing.Size(263, 22);
            this.dtpToDate.TabIndex = 1;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtpToDate);
            this.groupBox2.Location = new System.Drawing.Point(457, 255);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Size = new System.Drawing.Size(281, 53);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "To Date";
            // 
            // dtpLateTimeStart
            // 
            this.dtpLateTimeStart.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpLateTimeStart.Location = new System.Drawing.Point(55, 23);
            this.dtpLateTimeStart.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpLateTimeStart.Name = "dtpLateTimeStart";
            this.dtpLateTimeStart.ShowUpDown = true;
            this.dtpLateTimeStart.Size = new System.Drawing.Size(127, 22);
            this.dtpLateTimeStart.TabIndex = 1;
            this.dtpLateTimeStart.Value = new System.DateTime(2017, 2, 20, 8, 0, 0, 0);
            // 
            // gbxLateTimeRange
            // 
            this.gbxLateTimeRange.Controls.Add(this.label2);
            this.gbxLateTimeRange.Controls.Add(this.label1);
            this.gbxLateTimeRange.Controls.Add(this.dtpLateTimeEnd);
            this.gbxLateTimeRange.Controls.Add(this.dtpLateTimeStart);
            this.gbxLateTimeRange.Location = new System.Drawing.Point(19, 319);
            this.gbxLateTimeRange.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbxLateTimeRange.Name = "gbxLateTimeRange";
            this.gbxLateTimeRange.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbxLateTimeRange.Size = new System.Drawing.Size(391, 64);
            this.gbxLateTimeRange.TabIndex = 4;
            this.gbxLateTimeRange.TabStop = false;
            this.gbxLateTimeRange.Text = "Late Time Range";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(205, 28);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "End";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 28);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Start";
            // 
            // dtpLateTimeEnd
            // 
            this.dtpLateTimeEnd.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.dtpLateTimeEnd.Location = new System.Drawing.Point(248, 23);
            this.dtpLateTimeEnd.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpLateTimeEnd.Name = "dtpLateTimeEnd";
            this.dtpLateTimeEnd.ShowUpDown = true;
            this.dtpLateTimeEnd.Size = new System.Drawing.Size(127, 22);
            this.dtpLateTimeEnd.TabIndex = 2;
            this.dtpLateTimeEnd.Value = new System.DateTime(2017, 2, 20, 8, 0, 0, 0);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.tbxCnic);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.cbxCompany);
            this.groupBox4.Controls.Add(this.label10);
            this.groupBox4.Controls.Add(this.cbxCadre);
            this.groupBox4.Controls.Add(this.label9);
            this.groupBox4.Controls.Add(this.cbxCrew);
            this.groupBox4.Controls.Add(this.label8);
            this.groupBox4.Controls.Add(this.tbxCarNumber);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.tbxPNumber);
            this.groupBox4.Controls.Add(this.label6);
            this.groupBox4.Controls.Add(this.tbxName);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.cbxSections);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.cbxDepartments);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(19, 15);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox4.Size = new System.Drawing.Size(720, 233);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Filter by";
            // 
            // tbxCnic
            // 
            this.tbxCnic.Location = new System.Drawing.Point(105, 23);
            this.tbxCnic.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbxCnic.Name = "tbxCnic";
            this.tbxCnic.Size = new System.Drawing.Size(245, 22);
            this.tbxCnic.TabIndex = 21;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(8, 28);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(93, 17);
            this.label11.TabIndex = 20;
            this.label11.Text = "CNIC Number";
            // 
            // cbxCompany
            // 
            this.cbxCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCompany.FormattingEnabled = true;
            this.cbxCompany.Location = new System.Drawing.Point(460, 166);
            this.cbxCompany.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cbxCompany.Name = "cbxCompany";
            this.cbxCompany.Size = new System.Drawing.Size(249, 24);
            this.cbxCompany.TabIndex = 19;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(368, 173);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(67, 17);
            this.label10.TabIndex = 18;
            this.label10.Text = "Company";
           
            // 
            // cbxCadre
            // 
            this.cbxCadre.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCadre.FormattingEnabled = true;
            this.cbxCadre.Location = new System.Drawing.Point(459, 55);
            this.cbxCadre.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cbxCadre.Name = "cbxCadre";
            this.cbxCadre.Size = new System.Drawing.Size(249, 24);
            this.cbxCadre.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(368, 62);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(46, 17);
            this.label9.TabIndex = 16;
            this.label9.Text = "Cadre";
            // 
            // cbxCrew
            // 
            this.cbxCrew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxCrew.FormattingEnabled = true;
            this.cbxCrew.Location = new System.Drawing.Point(460, 93);
            this.cbxCrew.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cbxCrew.Name = "cbxCrew";
            this.cbxCrew.Size = new System.Drawing.Size(251, 24);
            this.cbxCrew.TabIndex = 15;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(368, 100);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(39, 17);
            this.label8.TabIndex = 14;
            this.label8.Text = "Crew";
            // 
            // tbxCarNumber
            // 
            this.tbxCarNumber.Location = new System.Drawing.Point(101, 114);
            this.tbxCarNumber.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbxCarNumber.Name = "tbxCarNumber";
            this.tbxCarNumber.Size = new System.Drawing.Size(249, 22);
            this.tbxCarNumber.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 118);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(60, 17);
            this.label7.TabIndex = 12;
            this.label7.Text = "Card No";
            // 
            // tbxPNumber
            // 
            this.tbxPNumber.Location = new System.Drawing.Point(101, 156);
            this.tbxPNumber.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbxPNumber.Name = "tbxPNumber";
            this.tbxPNumber.Size = new System.Drawing.Size(249, 22);
            this.tbxPNumber.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 161);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 17);
            this.label6.TabIndex = 10;
            this.label6.Text = "P-Number";
            // 
            // tbxName
            // 
            this.tbxName.Location = new System.Drawing.Point(105, 68);
            this.tbxName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbxName.Name = "tbxName";
            this.tbxName.Size = new System.Drawing.Size(245, 22);
            this.tbxName.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 73);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(45, 17);
            this.label5.TabIndex = 8;
            this.label5.Text = "Name";
            // 
            // cbxSections
            // 
            this.cbxSections.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxSections.FormattingEnabled = true;
            this.cbxSections.Location = new System.Drawing.Point(459, 16);
            this.cbxSections.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cbxSections.Name = "cbxSections";
            this.cbxSections.Size = new System.Drawing.Size(249, 24);
            this.cbxSections.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(368, 23);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 17);
            this.label4.TabIndex = 6;
            this.label4.Text = "Section";
           
            // 
            // cbxDepartments
            // 
            this.cbxDepartments.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbxDepartments.FormattingEnabled = true;
            this.cbxDepartments.Location = new System.Drawing.Point(459, 129);
            this.cbxDepartments.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cbxDepartments.Name = "cbxDepartments";
            this.cbxDepartments.Size = new System.Drawing.Size(251, 24);
            this.cbxDepartments.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(368, 136);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "Department";
            // 
            // Form1
            // 
            this.AcceptButton = this.btnGenerate;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(760, 394);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.gbxLateTimeRange);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnGenerate);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Engro Fertilizer";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.gbxLateTimeRange.ResumeLayout(false);
            this.gbxLateTimeRange.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        



        #endregion

        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.DateTimePicker dtpFromDate;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dtpToDate;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker dtpLateTimeStart;
        private System.Windows.Forms.GroupBox gbxLateTimeRange;
        private System.Windows.Forms.DateTimePicker dtpLateTimeEnd;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ComboBox cbxDepartments;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbxSections;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbxPNumber;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbxName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbxCarNumber;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cbxCrew;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbxCadre;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbxCnic;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbxCompany;
        private System.Windows.Forms.Label label10;
    }
}

