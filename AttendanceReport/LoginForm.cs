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

            List<Users> users = (from user in EFERTDbUtility.mEFERTDb.Users
                                 where user != null
                                 select user).ToList();

            Users loggedInUser = null;

            for (int i = 0; i < users.Count; i++)
            {
                Users user = users[i];
                string upasword = Helper.DecryptString(user.Password,
                        Helper.CONST_ENC_PASSPHRASE,
                        Helper.CONST_ENC_SALT_VALUE,
                        Helper.CONST_ENC_HASH_ALGO,
                        Helper.CONST_ENC_PASSWORD_ITERATION,
                        Helper.CONST_ENC_INIT_VECTOR,
                        Helper.CONST_ENC_KEY_SIZE);


                if (user.Name == userName && upasword == password)
                {

                    mLoggedInUser = new AppUser()
                    {
                        UserName = user.Name,
                        Password = password,
                        IsAdmin = user.Role == UserRole.Admin.ToString() ? true : false
                    };

                    loggedInUser = user;

                    break;
                }

            }



            if (loggedInUser == null)
            {
                MessageBox.Show(this, "Either username or password is incorrect.");
            }
            else if (string.IsNullOrEmpty(loggedInUser.Name) || string.IsNullOrEmpty(loggedInUser.Password))
            {
                MessageBox.Show(this, "Something is wrong please contact your database admin.");
            }
            else
            {

                DateTime result;
                bool updateDate = false;

                if (loggedInUser.LastLoginDate == null)
                {
                    bool parse = DateTime.TryParse(DateTime.Now.ToString(), out result);
                    if (parse)
                    {
                        updateDate = true;
                    }
                }
                else
                {
                    bool parse = DateTime.TryParse(loggedInUser.LastLoginDate.ToString(), out result);
                }


                if (!updateDate)
                {
                    TimeSpan t = DateTime.Now.Subtract(result);

                    if (t.TotalDays > 60)
                    {
                        MessageBox.Show("Your account freezed please contact your administrator.");
                        return;
                    }
                    else
                    {
                        if (DateTime.Now.Day != result.Day || DateTime.Now.Month != result.Month || DateTime.Now.Year != result.Year)
                        {
                            bool parse = DateTime.TryParse(DateTime.Now.ToString(), out result);
                            if (parse)
                            {
                                updateDate = true;
                            }
                        }
                    }
                }


                if (updateDate)
                {
                    try
                    {

                        loggedInUser.LastLoginDate = result;
                        EFERTDbUtility.mEFERTDb.Entry(loggedInUser).State = System.Data.Entity.EntityState.Modified;
                        EFERTDbUtility.mEFERTDb.SaveChanges();
                    }
                    catch (Exception ex)
                    {
                        //DebugLogger.LogInfo(ex);
                        EFERTDbUtility.RollBack();
                    }
                }

                ReportSelectorForm lsf = new ReportSelectorForm();

                lsf.Show();

                this.Hide();
            }

        }
    }
}
