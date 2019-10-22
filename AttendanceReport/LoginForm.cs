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
using Newtonsoft.Json;

namespace AttendanceReport
{
    public partial class LoginForm : Form
    {
        public static LoginForm mMainForm = null;
        public static AppUser mLoggedInUser = null;
      

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
                        UserId=user.UserId,
                        UserName = user.Name,
                        Password = password,
                        IsAdmin = user.Role == UserRole.Admin.ToString() ? true : false,
                        Role = user.Role
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

                UserInfoExtended userInfoExtended = null;

                if (!string.IsNullOrEmpty(loggedInUser.CustomData))
                {
                    userInfoExtended = JsonConvert.DeserializeObject<UserInfoExtended>(loggedInUser.CustomData);

                    if (userInfoExtended != null)
                    {

                        if (userInfoExtended.UserStatus == UserStatus.Disabled)
                        {
                            MessageBox.Show("Your account disabled please contact your administrator.");
                            return;
                        }
                    }
                }

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
                        if (loggedInUser.Role != UserRole.Admin.ToString())
                        {
                            if (userInfoExtended != null)
                            {
                                if (userInfoExtended.UserStatus != UserStatus.Disabled)
                                {
                                    try
                                    {
                                        userInfoExtended.UserStatus = UserStatus.Disabled;
                                        loggedInUser.CustomData = JsonConvert.SerializeObject(userInfoExtended);
                                        EFERTDbUtility.mEFERTDb.Entry(loggedInUser).State = System.Data.Entity.EntityState.Modified;
                                        EFERTDbUtility.mEFERTDb.SaveChanges();
                                    }
                                    catch (Exception ex)
                                    {
                                        EFERTDbUtility.RollBack();
                                    }
                                }
                            }

                            MessageBox.Show("Your account disabled because you have not logged in since last 60 days please contact your administrator.");
                            return;
                        }

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
                        if (loggedInUser.Role == "Normal")
                        {
                            loggedInUser.Role = UserRole.User.ToString();
                        }

                        if (string.IsNullOrEmpty(loggedInUser.CustomData))
                        {
                            UserInfoExtended extended = new UserInfoExtended();
                            extended.UserActiveDate = DateTime.Now;
                            extended.UserStatus = UserStatus.Enabled;
                            string str = JsonConvert.SerializeObject(extended);
                            loggedInUser.CustomData = str;
                        }
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

        private void LoginForm_Load(object sender, EventArgs e)
        {
        }
    }
}
