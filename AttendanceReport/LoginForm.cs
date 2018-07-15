using AttendanceReport.EFERTDb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class LoginForm : Form
    {
        public static LoginForm mMainForm = null;
        public static AppUser mLoggedInUser = null;

        private List<AppUser> mUsers = new List<AppUser>()
        {
            new AppUser()
            {
                UserName = "Admin",
                Password = "efert123#@!",
                IsAdmin = true
            },
            new AppUser()
            {
                UserName = "user",
                Password = "Engro786",
                IsAdmin = true
            },
            new AppUser()
            {
                UserName = "Guest1",
                Password = "Engro786",
                IsAdmin = true
            },
            new AppUser()
            {
                UserName = "Guest2",
                Password = "Engro786",
                IsAdmin = true
            }
        };

        public LoginForm()
        {
            InitializeComponent();
            this.lblVersion.Text = "Version: " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.tbxUserName.Select();
            mMainForm = this;
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string userName = this.tbxUserName.Text;
            string password = this.tbxPassword.Text;

            AppUser loggedInUser = (from user in mUsers
                                    where user != null && user.UserName == userName && user.Password == password
                                    select user).FirstOrDefault();

            if (loggedInUser == null)
            {
                MessageBox.Show(this, "Either username or password is incorrect.");
            }
            else
            {
                mLoggedInUser = loggedInUser;

                ReportSelectorForm lsf = new ReportSelectorForm();

                lsf.Show();

                this.Hide();
            }

        }
    }
}
