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
    public partial class AddManualData : Form
    {
        public AddManualData()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            int Manual_EffertNoofEmployee = (int)total_Errert_Employees.Value;
            int Manual_EffertHours = (int)total_Errert_Employees_Hours.Value;
            int Manual_OtherWorkers = (int)total_others_Employees.Value;
            int Manual_Othehours = (int)total_others_Employees_Hours.Value; ;

            Man_Hours_Report form = new Man_Hours_Report(Manual_EffertNoofEmployee,Manual_EffertHours,Manual_OtherWorkers,Manual_Othehours);

            form.Show(this);
        }
    }
}
