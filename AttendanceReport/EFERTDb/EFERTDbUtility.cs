using AttendanceReport.CCFTCentral;
using AttendanceReport.EFERTDb;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public static class EFERTDbUtility
    {
        public static CCFTCentral.CCFTCentral mCCFTCentral = null;
        public static CCFTEvent.CCFTEvent mCCFTEvent = null;
        public static EFERTDbContext mEFERTDb = null;
        public static List<VisitingLocations> mVisitingLocations = null;

        public static string CONST_SYSTEM_BLOCKED_BY = "System";
        public static string CONST_SYSTEM_LIMIT_REACHED_REASON = "System has block this person because maximum limit of temporary check in has reached.";

        public static void InitializeDatabases()
        {
            mEFERTDb = new EFERTDbContext();
            mCCFTCentral = new CCFTCentral.CCFTCentral();
            mCCFTEvent = new CCFTEvent.CCFTEvent();

            try
            {
                Cardholder cardHolderByNic = (from pds in mCCFTCentral.PersonalDataStrings
                                              where pds != null && pds.PersonalDataFieldID == 5051 && pds.Value != null && pds.Value == "12345-1234567-1"
                                              select pds.Cardholder).FirstOrDefault();
            }
            catch (Exception e)
            {
                
            }

            try
            {
                mVisitingLocations = (from visitLocation in mEFERTDb.VisitingLocations
                                      where visitLocation != null
                                      select visitLocation).ToList();

                 List<Users> users = (from user1 in EFERTDbUtility.mEFERTDb.Users
                                       where user1 != null
                                       select user1).ToList();
                if (users == null || users.Count == 0)
                {
                    string userPasword = Helper.EncryptString("efert123#@!",
                  Helper.CONST_ENC_PASSPHRASE,
                  Helper.CONST_ENC_SALT_VALUE,
                  Helper.CONST_ENC_HASH_ALGO,
                  Helper.CONST_ENC_PASSWORD_ITERATION,
                  Helper.CONST_ENC_INIT_VECTOR,
                  Helper.CONST_ENC_KEY_SIZE);

                    Users user = new Users()
                    {
                        Name = "Admin",
                        Password = userPasword,
                        Role = "Admin"
                    };
                    EFERTDbUtility.mEFERTDb.Users.Add(user);
                    EFERTDbUtility.mEFERTDb.SaveChanges();
                }
            }
            catch (Exception e)
            {
                
            }

            try
            {
                List<CCFTEvent.Event> lstEvents = (from events in mCCFTEvent.Events
                                                   where events != null && events.EventType == 20001
                                                   select events).ToList();
            }
            catch (Exception e)
            {

            }
            
        }

        public static void RollBack()
        {
            var context = mEFERTDb;
            var changedEntries = context.ChangeTracker.Entries()
                .Where(x => x.State != EntityState.Unchanged).ToList();

            foreach (var entry in changedEntries)
            {
                switch (entry.State)
                {
                    case EntityState.Modified:
                        entry.CurrentValues.SetValues(entry.OriginalValues);
                        entry.State = EntityState.Unchanged;
                        break;
                    case EntityState.Added:
                        entry.State = EntityState.Detached;
                        break;
                    case EntityState.Deleted:
                        entry.State = EntityState.Unchanged;
                        break;
                }
            }
        }

        public static string GetInnerExceptionMessage(Exception ex)
        {
            string message = ex.Message;
            Exception innerException = ex.InnerException;

            while (innerException != null)
            {
                message += "\n" + innerException.Message;
                innerException = innerException.InnerException;
            }

            return message;
        }
                
        public static LimitStatus CheckIfUserCheckedInLimitReached(List<CheckInAndOutInfo> checkIns, List<BlockedPersonInfo> blocks, ref int refCheckinCount, bool sendEmail = true)
        {
            LimitStatus limitStatus = LimitStatus.Allowed;
            SystemSetting setting = EFERTDbUtility.mEFERTDb.SystemSetting.FirstOrDefault();

            int daysToEmailNotification = setting == null ? 70 : setting.DaysToEmailNotification;
            int daysToBlock = setting == null ? 90 : setting.DaysToBlockUser;

            if (checkIns.Count > 0)
            {
                CheckInAndOutInfo last = checkIns.Last();
                   
                if (last.CheckedIn)
                {
                    limitStatus = LimitStatus.CurrentlyCheckIn;
                }
                else
                {
                    //system deplymenrt date is 1/jan/2017 requirement is get user yearly report
                    DateTime fromDate = new DateTime(2018, 04, 1);
                    DateTime toDate = DateTime.Now;

                    BlockedPersonInfo lastBlockedPerson = (from block in blocks
                                                           where block != null &&
                                                                 block.BlockedTime >= fromDate &&
                                                                 block.BlockedTime < toDate && 
                                                                 !block.Blocked && 
                                                                 block.BlockedBy == CONST_SYSTEM_BLOCKED_BY
                                                           select block).LastOrDefault();

                    if (lastBlockedPerson != null)
                    {
                        fromDate = lastBlockedPerson.UnBlockTime;
                    }

                    checkIns = (from checkin in checkIns
                                where checkin != null && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate
                                select checkin).ToList();

                    if (checkIns.Count == 0)
                    {
                        return limitStatus;
                    }

                    string name, cnic = string.Empty;

                    if (last.CardHolderInfos != null)
                    {
                        name = last.CardHolderInfos.FirstName;
                        cnic = last.CardHolderInfos.CNICNumber;
                    }
                    else if (last.Visitors != null)
                    {
                        name = last.Visitors.FirstName;
                        cnic = last.Visitors.CNICNumber;
                    }
                    else
                    {
                        name = last.DailyCardHolders.FirstName;
                        cnic = last.DailyCardHolders.CNICNumber;
                    }

                    List<AlertInfo> chAlertInfos = (from alert in EFERTDbUtility.mEFERTDb.AlertInfos
                                                    where alert != null && alert.CNICNumber == cnic
                                                    select alert).ToList();

                    DateTime alertEnableDate = DateTime.MaxValue;
                    bool alertEnabled = true;
                    AlertInfo lastAlertInfo = null;

                    if (chAlertInfos != null && chAlertInfos.Count > 0)
                    {
                        lastAlertInfo = chAlertInfos.Last();

                        if (lastAlertInfo.DisableAlert)
                        {
                            alertEnabled = false;
                        }
                        else
                        {
                            alertEnableDate = lastAlertInfo.EnableAlertDate;
                        }
                    }


                    CheckInAndOutInfo previousCheckIn = checkIns[0];
                    bool isSameDay = last.DateTimeIn.Date == DateTime.Now.Date;

                    if (previousCheckIn != null)
                    {
                        int count = 1;
                        DateTime previousDateTimeIn = previousCheckIn.DateTimeIn;

                        for (int i = 1; i < checkIns.Count; i++)
                        {
                            CheckInAndOutInfo CurrentCheckIn = checkIns[i];

                            DateTime currDateTimeIn = CurrentCheckIn.DateTimeIn;

                            TimeSpan timeDiff = currDateTimeIn.Date - previousDateTimeIn.Date;



                            //bool isContinous = timeDiff.Days == 1 || timeDiff.Days == 2 && currDateTimeIn.DayOfWeek == DayOfWeek.Monday;

                            if (timeDiff.Days >= 1 )
                            {
                                count++;
                            }
                            //else
                            //{
                            //    if (currDateTimeIn.Date != previousDateTimeIn.Date)
                            //    {
                            //        count = 1;
                            //    }

                            //}

                            previousDateTimeIn = currDateTimeIn;
                        }

                        refCheckinCount = count;

                        if (count >= daysToEmailNotification)
                        {
                            if (count == daysToBlock && !isSameDay)
                            {
                                limitStatus = LimitStatus.LimitReached;
                            }
                            else
                            {
                                if (alertEnabled)
                                {
                                    if (sendEmail)
                                    {
                                        List<EmailAddress> toAddresses = new List<EmailAddress>();

                                        if (EFERTDbUtility.mEFERTDb.EmailAddresses != null)
                                        {
                                            toAddresses = (from email in EFERTDbUtility.mEFERTDb.EmailAddresses
                                                           where email != null
                                                           select email).ToList();

                                            foreach (EmailAddress toAddress in toAddresses)
                                            {
                                                SendMail(setting, toAddress.Email, toAddress.Name, name, cnic);
                                            }
                                        }
                                    }


                                    limitStatus = LimitStatus.EmailAlerted;
                                }
                                else
                                {
                                    limitStatus = LimitStatus.EmailAlertDisabled;
                                }
                            }

                        }
                        else
                        {
                            limitStatus = LimitStatus.Allowed;
                        }
                    }
                }
            }

            return limitStatus;
        }

        public static bool ValidateInputs(List<TextBox> lstTextBoxes)
        {
            bool validated = false;

            if (lstTextBoxes != null)
            {
                List<string> lstInvalidTextboxes = (from text in lstTextBoxes
                 where text != null && string.IsNullOrEmpty(text.Text)
                 select text.Name).ToList();

                validated = lstInvalidTextboxes.Count == 0;

                foreach (TextBox textbox in lstTextBoxes)
                {
                    if (textbox.ReadOnly)
                    {
                        continue;
                    }
                    if (lstInvalidTextboxes.Contains(textbox.Name))
                    {
                        textbox.BackColor = System.Drawing.Color.Yellow;
                    }
                    else
                    {
                        textbox.BackColor = System.Drawing.Color.White;
                    }

                }
            }

            return validated;
        }

        public static byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            byte[] array = null;

            if (imageIn != null)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
                    array = ms.ToArray();
                }
            }

            return array;
        }

        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            Image returnImage = null;

            if (byteArrayIn != null)
            {
                using (MemoryStream ms = new MemoryStream(byteArrayIn))
                {
                    returnImage = Image.FromStream(ms);
                }
            }
            
            return returnImage;
        }

        public static void AllowNumericOnly(System.Windows.Forms.KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

        public static void SendMail(SystemSetting settings, string toAddress, string toName, string chName, string chCnic)
        {
            try
            {
                string fromaddr = settings.FromEmailAddress;
                //string toAddress = "hassanwebdeveloper30@gmail.com";//TO ADDRESS HERE
                string password = settings.FromEmailPassword;

                int emailNotificationDays = settings.DaysToEmailNotification;

                using (MailMessage msg = new MailMessage())
                {

                    msg.Subject = "Security Alert for continously entry of casual worker.";

                    msg.From = new MailAddress(fromaddr);
                    msg.Body = "Dear Sir,\n\nIt's for your information following worker entry limit reached at " + emailNotificationDays + " day(s) need your necessary action on it.\n\n Name: "+chName+"\n\n CNIC Number: "+chCnic+"\n\nThis is system generated email.";
                    msg.To.Add(new MailAddress(toAddress));

                    using (SmtpClient smtp = new SmtpClient())
                    {
                        smtp.Host = settings.SmtpServer;
                        smtp.Port = Convert.ToInt32(settings.SmtpPort);
                        smtp.UseDefaultCredentials = true;
                        smtp.EnableSsl = settings.IsSmptSSL;

                        if (settings.IsSmptAuthRequired)
                        {
                            NetworkCredential nc = new NetworkCredential(fromaddr, password);
                            smtp.Credentials = nc;
                        }

                        smtp.Send(msg);
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured in sending email.\n\n" + GetInnerExceptionMessage(ex));
            }
        }

        public static void UpdateDropDownFields(ComboBox cbxDepartments, ComboBox cbxSections, ComboBox cbxCompany, ComboBox cbxCadre, ComboBox cbxCrew)
        {
            if (cbxDepartments != null)
            {
                List<string> ccftDepartments = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                            select pds.Value).Distinct().ToList();

                List<string> departments = (from depart in EFERTDbUtility.mEFERTDb.Departments
                                            where depart != null && !string.IsNullOrEmpty(depart.DepartmentName)
                                            select depart.DepartmentName).ToList();

                departments.Insert(0, string.Empty);

                departments = departments.TakeWhile(a => !ccftDepartments.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftDepartments.AddRange(departments);

                ccftDepartments.Sort();

                cbxDepartments.Items.AddRange(ccftDepartments.ToArray());
            }

            if (cbxSections != null)
            {
                List<string> ccftSections = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                         where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                         select pds.Value).Distinct().ToList();

                List<string> sections = (from section in EFERTDbUtility.mEFERTDb.Sections
                                         where section != null && !string.IsNullOrEmpty(section.SectionName)
                                         select section.SectionName).ToList();

                sections.Insert(0, string.Empty);

                sections = sections.TakeWhile(a => !ccftSections.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftSections.AddRange(sections);

                ccftSections.Sort();

                cbxSections.Items.AddRange(ccftSections.ToArray());
            }

            if (cbxCompany != null)
            {
                List<string> ccftCompanyNames = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                             where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                             select pds.Value).Distinct().ToList();

                List<string> companies = (from company in EFERTDbUtility.mEFERTDb.Companies
                                          where company != null && !string.IsNullOrEmpty(company.CompanyName)
                                          select company.CompanyName).ToList();

                companies.Insert(0, string.Empty);

                companies = companies.TakeWhile(a => !ccftCompanyNames.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftCompanyNames.AddRange(companies);

                ccftCompanyNames.Sort();

                cbxCompany.Items.AddRange(ccftCompanyNames.ToArray());
            }

            if (cbxCadre != null)
            {
                List<string> ccftCadres = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                       select pds.Value).Distinct().ToList();

                List<string> cadres = (from cadre in EFERTDbUtility.mEFERTDb.Cadres
                                       where cadre != null && !string.IsNullOrEmpty(cadre.CadreName)
                                       select cadre.CadreName).ToList();

                cadres.Insert(0, string.Empty);

                cadres = cadres.TakeWhile(a => !ccftCadres.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftCadres.AddRange(cadres);

                ccftCadres.Sort();

                cbxCadre.Items.AddRange(ccftCadres.ToArray());
            }

            if (cbxCrew != null)
            {
                List<string> crews = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                      where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                      select pds.Value).Distinct().ToList();

                crews.Insert(0, string.Empty);
                
                crews.Sort();

                cbxCrew.Items.AddRange(crews.ToArray());

            }

        }


        public static void UpdateDropDownFieldsMultiCombobox(CheckedComboBox cbxDepartments, CheckedComboBox cbxSections, CheckedComboBox cbxCompany, CheckedComboBox cbxCadre, CheckedComboBox cbxCrew)
        {
            if (cbxDepartments != null)
            {
                List<string> ccftDepartments = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                                where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                                select pds.Value).Distinct().ToList();

                List<string> departments = (from depart in EFERTDbUtility.mEFERTDb.Departments
                                            where depart != null && !string.IsNullOrEmpty(depart.DepartmentName)
                                            select depart.DepartmentName).ToList();

                

                departments = departments.TakeWhile(a => !ccftDepartments.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftDepartments.AddRange(departments);

                ccftDepartments.Sort();

                CCBoxItem items = new CCBoxItem("UnChek All", 0);
                cbxDepartments.Items.Add(items);
                for (int i = 0; i < ccftDepartments.Count; i++)
                {
                    CCBoxItem item = new CCBoxItem(ccftDepartments[i], i + 1);
                    cbxDepartments.Items.Add(item);
                }
                // If more then 5 items, add a scroll bar to the dropdown.
                cbxDepartments.MaxDropDownItems = 20;
                cbxDepartments.DisplayMember = "Name";

            }

            if (cbxSections != null)
            {
                List<string> ccftSections = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                             where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                             select pds.Value).Distinct().ToList();

                List<string> sections = (from section in EFERTDbUtility.mEFERTDb.Sections
                                         where section != null && !string.IsNullOrEmpty(section.SectionName)
                                         select section.SectionName).ToList();

                sections.Insert(0, string.Empty);

                sections = sections.TakeWhile(a => !ccftSections.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftSections.AddRange(sections);

                ccftSections.Sort();

                CCBoxItem items = new CCBoxItem("UnChek All", 0);
                cbxSections.Items.Add(items);
                for (int i = 0; i < ccftSections.Count; i++)
                {
                    CCBoxItem item = new CCBoxItem(ccftSections[i], i + 1);
                    cbxSections.Items.Add(item);
                }
                // If more then 5 items, add a scroll bar to the dropdown.
                cbxSections.MaxDropDownItems = 20;
                cbxSections.DisplayMember = "Name";
            }

            if (cbxCompany != null)
            {
                List<string> ccftCompanyNames = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                                 where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                                 select pds.Value).Distinct().ToList();

                List<string> companies = (from company in EFERTDbUtility.mEFERTDb.Companies
                                          where company != null && !string.IsNullOrEmpty(company.CompanyName)
                                          select company.CompanyName).ToList();



                companies = companies.TakeWhile(a => !ccftCompanyNames.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftCompanyNames.AddRange(companies);

                ccftCompanyNames.Sort();

                CCBoxItem items = new CCBoxItem("UnChek All", 0);
                cbxCompany.Items.Add(items);
                for (int i = 0; i < ccftCompanyNames.Count; i++)
                {
                    CCBoxItem item = new CCBoxItem(ccftCompanyNames[i], i + 1);
                    cbxCompany.Items.Add(item);

                }
                // If more then 5 items, add a scroll bar to the dropdown.
                cbxCompany.MaxDropDownItems = 20;
                cbxCompany.DisplayMember = "Name";
            }

            if (cbxCadre != null)
            {
                List<string> ccftCadres = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                           where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                           select pds.Value).Distinct().ToList();

                List<string> cadres = (from cadre in EFERTDbUtility.mEFERTDb.Cadres
                                       where cadre != null && !string.IsNullOrEmpty(cadre.CadreName)
                                       select cadre.CadreName).ToList();

                cadres.Insert(0, string.Empty);

                cadres = cadres.TakeWhile(a => !ccftCadres.Exists(b => b.ToLower() == a.ToLower())).ToList();

                ccftCadres.AddRange(cadres);

                ccftCadres.Sort();

                CCBoxItem items = new CCBoxItem("UnChek All", 0);
                cbxCadre.Items.Add(items);
                for (int i = 0; i < ccftCadres.Count; i++)
                {
                    CCBoxItem item = new CCBoxItem(ccftCadres[i], i + 1);
                    cbxCadre.Items.Add(item);
                }
                // If more then 5 items, add a scroll bar to the dropdown.
                cbxCadre.MaxDropDownItems = 20;
                cbxCadre.DisplayMember = "Name";
            }

            if (cbxCrew != null)
            {
                List<string> crews = (from pds in EFERTDbUtility.mCCFTCentral.PersonalDataStrings
                                      where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                      select pds.Value).Distinct().ToList();

                crews.Insert(0, string.Empty);

                crews.Sort();

                CCBoxItem items = new CCBoxItem("UnChek All", 0);
                cbxCrew.Items.Add(items);
                for (int i = 0; i < crews.Count; i++)
                {
                    CCBoxItem item = new CCBoxItem(crews[i], i + 1);
                    cbxCrew.Items.Add(item);

                }
                // If more then 5 items, add a scroll bar to the dropdown.
                cbxCrew.MaxDropDownItems = 20;
                cbxCrew.DisplayMember = "Name";

            }

        }

    }




    public enum LimitStatus
    {
        Allowed,
        LimitReached,
        CurrentlyCheckIn,
        EmailAlerted,
        EmailAlertDisabled
    }
}
