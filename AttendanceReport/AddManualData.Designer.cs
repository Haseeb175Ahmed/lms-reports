
namespace AttendanceReport
{
    partial class AddManualData
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
            this.total_others_Employees_Hours = new System.Windows.Forms.NumericUpDown();
            this.total_others_Employees = new System.Windows.Forms.NumericUpDown();
            this.total_Errert_Employees_Hours = new System.Windows.Forms.NumericUpDown();
            this.total_Errert_Employees = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.total_others_Employees_Hours)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_others_Employees)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_Errert_Employees_Hours)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_Errert_Employees)).BeginInit();
            this.SuspendLayout();
            // 
            // total_others_Employees_Hours
            // 
            this.total_others_Employees_Hours.Location = new System.Drawing.Point(236, 148);
            this.total_others_Employees_Hours.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.total_others_Employees_Hours.Name = "total_others_Employees_Hours";
            this.total_others_Employees_Hours.Size = new System.Drawing.Size(106, 20);
            this.total_others_Employees_Hours.TabIndex = 21;
            // 
            // total_others_Employees
            // 
            this.total_others_Employees.Location = new System.Drawing.Point(236, 121);
            this.total_others_Employees.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.total_others_Employees.Name = "total_others_Employees";
            this.total_others_Employees.Size = new System.Drawing.Size(106, 20);
            this.total_others_Employees.TabIndex = 20;
            // 
            // total_Errert_Employees_Hours
            // 
            this.total_Errert_Employees_Hours.Location = new System.Drawing.Point(236, 93);
            this.total_Errert_Employees_Hours.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.total_Errert_Employees_Hours.Name = "total_Errert_Employees_Hours";
            this.total_Errert_Employees_Hours.Size = new System.Drawing.Size(106, 20);
            this.total_Errert_Employees_Hours.TabIndex = 19;
            // 
            // total_Errert_Employees
            // 
            this.total_Errert_Employees.Location = new System.Drawing.Point(236, 67);
            this.total_Errert_Employees.Maximum = new decimal(new int[] {
            10000,
            0,
            0,
            0});
            this.total_Errert_Employees.Name = "total_Errert_Employees";
            this.total_Errert_Employees.Size = new System.Drawing.Size(106, 20);
            this.total_Errert_Employees.TabIndex = 18;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(144, 191);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(97, 34);
            this.button1.TabIndex = 17;
            this.button1.Text = "Next";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(23, 153);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(176, 15);
            this.label4.TabIndex = 16;
            this.label4.Text = "Other Workers Total Hours";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(23, 126);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(170, 15);
            this.label3.TabIndex = 15;
            this.label3.Text = "Other Workers Total Entry";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(23, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(167, 15);
            this.label2.TabIndex = 14;
            this.label2.Text = "EFERT Employees Hours";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(23, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(161, 15);
            this.label1.TabIndex = 13;
            this.label1.Text = "EFERT Employees Total";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(54, 20);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(288, 20);
            this.label5.TabIndex = 22;
            this.label5.Text = "Fill Custom Fields for Manual Entry";
            // 
            // AddManualData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 260);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.total_others_Employees_Hours);
            this.Controls.Add(this.total_others_Employees);
            this.Controls.Add(this.total_Errert_Employees_Hours);
            this.Controls.Add(this.total_Errert_Employees);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "AddManualData";
            this.Text = "AddManualData";
            ((System.ComponentModel.ISupportInitialize)(this.total_others_Employees_Hours)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_others_Employees)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_Errert_Employees_Hours)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.total_Errert_Employees)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NumericUpDown total_others_Employees_Hours;
        private System.Windows.Forms.NumericUpDown total_others_Employees;
        private System.Windows.Forms.NumericUpDown total_Errert_Employees_Hours;
        private System.Windows.Forms.NumericUpDown total_Errert_Employees;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
    }
}