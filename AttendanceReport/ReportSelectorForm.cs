using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class ReportSelectorForm : Form
    {
        public ReportSelectorForm()
        {
            InitializeComponent();

            this.UpdateLayout();
        }

        private void UpdateLayout()
        {
            if (LoginForm.mLoggedInUser.UserName == "user" || LoginForm.mLoggedInUser.UserName == "Guest1" || LoginForm.mLoggedInUser.UserName == "Guest2")
            {
                this.btnLateArrivalReport.Visible = false;
                this.btnAttendance.Visible = false;
                this.btnContractorWise.Visible = false;
                this.button1.Visible = false;
                this.btnCategoryWise.Visible = false;
                this.btnCategoryWiseNotReturned.Visible = false;
                this.btnBannedEntry.Visible = false;
                this.btnAlarmedListReport.Visible = false;
                this.btnAttendanceReport.Location = new Point(78, 39);
                this.Size = new Size(287, 169);
            }
        }

        private void btnLateArrivalReport_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1();

            form.Show(this);
        }

        private void ReportSelectorForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!LoginForm.mMainForm.Visible)
            {
                LoginForm.mMainForm.Close();
            }
        }

        private void btnAttendance_Click(object sender, EventArgs e)
        {
            ActivityReportForm form = new ActivityReportForm();

            form.Show(this);
        }

        private void btnContractorWise_Click(object sender, EventArgs e)
        {
            ContractorWiseReportForm form = new ContractorWiseReportForm(false);

            form.Show(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ContractorWiseReportForm form = new ContractorWiseReportForm(true);

            form.Show(this);
        }

        private void btnCategoryWise_Click(object sender, EventArgs e)
        {
            CategoryWiseReportForm form = new CategoryWiseReportForm(false);

            form.Show(this);
        }

        private void btnCategoryWiseNotReturned_Click(object sender, EventArgs e)
        {
            CategoryWiseReportForm form = new CategoryWiseReportForm(true);

            form.Show(this);
        }

        private void btnBannedEntry_Click(object sender, EventArgs e)
        {
            BannedEntryReportForm form = new BannedEntryReportForm();

            form.Show(this);
        }

        private void btnAlarmedListReport_Click(object sender, EventArgs e)
        {
            AlarmedListReportForm form = new AlarmedListReportForm();

            form.Show(this);
        }

        private void btnAttendanceReport_Click(object sender, EventArgs e)
        {
            LMSAttendanceReportForm form = new LMSAttendanceReportForm();

            form.Show(this);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddManualData form = new AddManualData();
            form.Show();
        }
    }
}
