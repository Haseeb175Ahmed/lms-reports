using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AttendanceReport.EFERTDb;
using AttendanceReport.CCFTEvent;
using AttendanceReport.CCFTCentral;
using System.IO;

namespace AttendanceReport
{
    public partial class Man_Hours_Report : Form
    {
        int Manual_EffertNoofEmployee = 0;
        int Manual_EffertHours = 0;
        int Manual_OtherWorkers = 0;
        int Manual_Othehours = 0;

        private Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> mData = null;

        Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> Depart_Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();

        Dictionary<string, Dictionary<string, Dictionary<string, double>>> Summary_Report = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
        

        //List<KeyValuePair<string, List<KeyValuePair<string, double>>>> DepartWise1 = new List<KeyValuePair<string, List<KeyValuePair<string, double>>>>();

        public Man_Hours_Report()
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFields(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre, null);


            //dtpFromDate.Value = Convert.ToDateTime("1/1/2020");
            //dtpToDate.Value = Convert.ToDateTime("1/10/2020");
        }

        public Man_Hours_Report(int p_Manual_EffertNoofEmployee, int p_Manual_EffertHours, int p_Manual_OtherWorkers, int p_Manual_Othehours)
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFields(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre, null);
            
            this.Manual_EffertNoofEmployee = p_Manual_EffertNoofEmployee;
            this.Manual_EffertHours = p_Manual_EffertHours;
            this.Manual_OtherWorkers = p_Manual_OtherWorkers;
            this.Manual_Othehours = p_Manual_Othehours;

            //dtpFromDate.Value = Convert.ToDateTime("1/1/2020");
            //dtpToDate.Value = Convert.ToDateTime("1/10/2020");
        }


        private async  void btnGenerate_Click(object sender, EventArgs e)
        {



            //create data object and print report.
            // Man - Hours Summary Report

            Depart_Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
            Summary_Report = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

            if (reportType.Text.ToString() == "Man-Hours Detail Report")
            {
                if (cbxDepartments.Text.ToString().ToUpper() == "EFERTDHKALL")
                {
                    MessageBox.Show("This Report is Not for All Departments.");
                }
                else
                {

                    await Task.Run(() =>
                    {
                        Generate_ManHours_Detail_report();
                    });

                    if (this.Depart_Date_CadreNic != null && this.Depart_Date_CadreNic.Count > 0)
                    {
                        //Cursor.Current = currentCursor;
                        this.saveFileDialog1.ShowDialog(this);

                    }
                    else
                    {
                        //Cursor.Current = currentCursor;
                        MessageBox.Show(this, "No data exist on current selected date range.");
                    }

                }
            }
            else
            {

                if (cbxDepartments.Text.ToString().ToUpper() != "EFERTDHKALL")
                {
                    MessageBox.Show("This Report is Not for Specific Department.");
                }
                else
                {
                    await Task.Run(() =>
                    {
                        Generate_ManHours_Summary_report();
                    });

                    if (this.Summary_Report != null && this.Summary_Report.Count > 0)
                    {
                        //Cursor.Current = currentCursor;
                        this.saveFileDialog2.ShowDialog(this);

                    }
                    else
                    {
                        //Cursor.Current = currentCursor;
                        MessageBox.Show(this, "No data exist on current selected date range.");
                    }
                }
            }
        }
        public void Generate_ManHours_Summary_report()
        {
            try
            {

                
                int progress = 0;


                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = true; }));
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }


                #region

                this.Summary_Report = null;
                DateTime fromDate = this.dtpFromDate.Value.Date;
                DateTime fromDateUtc = fromDate.ToUniversalTime();
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                DateTime toDateUtc = toDate.ToUniversalTime();

                DateTime ndtStartDate = this.dtpNdtStart.Value;
                DateTime ndtEndDate = this.dtpNdtEnd.Value;
                DateTime ndtLunchStartDate = this.dtpNdtLunchStart.Value;
                DateTime ndtLunchEndDate = this.dtpNdtLunchEnd.Value;

                DateTime fdtStartDate = this.dtpFdtStart.Value;
                DateTime fdtEndDate = this.dtpFdtEnd.Value;
                DateTime fdtLunchStartDate = this.dtpFdtLunchStart.Value;
                DateTime fdtLunchEndDate = this.dtpFdtLunchEnd.Value;

                int ndtGraceTimeBeforeStart = Convert.ToInt32(nudNdtGraceTimeBeforeStart.Value);
                int ndtGraceTimeAfterStart = Convert.ToInt32(nudNdtGraceTimeBeforeStart.Value);
                int ndtGraceTimeBeforeEnd = Convert.ToInt32(nudNdtGraceTimeBeforeEnd.Value);
                int ndtGraceTimeAfterEnd = Convert.ToInt32(nudNdtGraceTimeBeforeEnd.Value);
                int ndtGraceTimeBeforeLunchStart = Convert.ToInt32(nudNdtGraceTimeBeforeLunchStart.Value);
                int ndtGraceTimeAfterLunchStart = Convert.ToInt32(nudNdtGraceTimeBeforeLunchStart.Value);
                int ndtGraceTimeBeforeLunchEnd = Convert.ToInt32(nudNdtGraceTimeBeforeLunchEnd.Value);
                int ndtGraceTimeAfterLunchEnd = Convert.ToInt32(nudNdtGraceTimeBeforeLunchEnd.Value);

                int fdtGraceTimeBeforeStart = Convert.ToInt32(nudFdtGraceTimeBeforeStart.Value);
                int fdtGraceTimeAfterStart = Convert.ToInt32(nudFdtGraceTimeBeforeStart.Value);
                int fdtGraceTimeBeforeEnd = Convert.ToInt32(nudFdtGraceTimeBeforeEnd.Value);
                int fdtGraceTimeAfterEnd = Convert.ToInt32(nudFdtGraceTimeBeforeEnd.Value);
                int fdtGraceTimeBeforeLunchStart = Convert.ToInt32(nudFdtGraceTimeBeforeLunchStart.Value);
                int fdtGraceTimeAfterLunchStart = Convert.ToInt32(nudFdtGraceTimeBeforeLunchStart.Value);
                int fdtGraceTimeBeforeLunchEnd = Convert.ToInt32(nudFdtGraceTimeBeforeLunchEnd.Value);
                int fdtGraceTimeAfterLunchEnd = Convert.ToInt32(nudFdtGraceTimeBeforeLunchEnd.Value);

                TimeSpan ndtStartTime = this.dtpNdtStart.Value.TimeOfDay;
                TimeSpan ndtEndTime = this.dtpNdtEnd.Value.TimeOfDay;
                TimeSpan ndtLunchStartTime = this.dtpNdtLunchStart.Value.TimeOfDay;
                TimeSpan ndtLunchEndTime = this.dtpNdtLunchEnd.Value.TimeOfDay;

                TimeSpan fdtStartTime = this.dtpFdtStart.Value.TimeOfDay;
                TimeSpan fdtEndTime = this.dtpFdtEnd.Value.TimeOfDay;
                TimeSpan fdtLunchStartTime = this.dtpFdtLunchStart.Value.TimeOfDay;
                TimeSpan fdtLunchEndTime = this.dtpFdtLunchEnd.Value.TimeOfDay;

                TimeSpan ndtWithBeforeGraceTimeStartTime = this.dtpNdtStart.Value.AddMinutes(ndtGraceTimeBeforeStart * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeEndTime = this.dtpNdtEnd.Value.AddMinutes(ndtGraceTimeBeforeEnd * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeLunchStartTime = this.dtpNdtLunchStart.Value.AddMinutes(ndtGraceTimeBeforeLunchStart * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeLunchEndTime = this.dtpNdtLunchEnd.Value.AddMinutes(ndtGraceTimeBeforeLunchEnd * -1).TimeOfDay;

                TimeSpan ndtWithAfterGraceTimeStartTime = this.dtpNdtStart.Value.AddMinutes(ndtGraceTimeAfterStart).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeEndTime = this.dtpNdtEnd.Value.AddMinutes(ndtGraceTimeAfterEnd).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeLunchStartTime = this.dtpNdtLunchStart.Value.AddMinutes(ndtGraceTimeAfterLunchStart).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeLunchEndTime = this.dtpNdtLunchEnd.Value.AddMinutes(ndtGraceTimeAfterLunchEnd).TimeOfDay;

                TimeSpan fdtWithBeforeGraceTimeStartTime = this.dtpFdtStart.Value.AddMinutes(fdtGraceTimeBeforeStart * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeEndTime = this.dtpFdtEnd.Value.AddMinutes(fdtGraceTimeBeforeEnd * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeLunchStartTime = this.dtpFdtLunchStart.Value.AddMinutes(fdtGraceTimeBeforeLunchStart * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeLunchEndTime = this.dtpFdtLunchEnd.Value.AddMinutes(fdtGraceTimeBeforeLunchEnd * -1).TimeOfDay;

                TimeSpan fdtWithAfterGraceTimeStartTime = this.dtpFdtStart.Value.AddMinutes(fdtGraceTimeAfterStart).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeEndTime = this.dtpFdtEnd.Value.AddMinutes(fdtGraceTimeAfterEnd).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeLunchStartTime = this.dtpFdtLunchStart.Value.AddMinutes(fdtGraceTimeAfterLunchStart).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeLunchEndTime = this.dtpFdtLunchEnd.Value.AddMinutes(fdtGraceTimeAfterLunchEnd).TimeOfDay;

                string filterByDepartment = "";
                string filterBySection = "";
                string filerByName = "";
                string filterByCadre = "";
                string filterByCompany = "";
                string filterByCNIC = "";
                string filterByPnumber = "";

                if (cbxDepartments.InvokeRequired)
                {
                    cbxDepartments.Invoke(new MethodInvoker(delegate { filterByDepartment = this.cbxDepartments.Text.ToLower(); }));
                }

                if (cbxSections.InvokeRequired)
                {
                    cbxSections.Invoke(new MethodInvoker(delegate { filterBySection = this.cbxSections.Text.ToLower(); }));
                }


                if (tbxName.InvokeRequired)
                {
                    tbxName.Invoke(new MethodInvoker(delegate { filerByName = this.tbxName.Text.ToLower(); }));
                }


                if (cbxCadre.InvokeRequired)
                {
                    cbxCadre.Invoke(new MethodInvoker(delegate { filterByCadre = this.cbxCadre.Text.ToLower(); }));
                }


                if (cbxCompany.InvokeRequired)
                {
                    cbxCompany.Invoke(new MethodInvoker(delegate { filterByCompany = this.cbxCompany.Text.ToLower(); }));
                }


                if (tbxCnic.InvokeRequired)
                {
                    tbxCnic.Invoke(new MethodInvoker(delegate { filterByCNIC = this.tbxCnic.Text; }));
                }


                if (tbxPNumber.InvokeRequired)
                {
                    tbxPNumber.Invoke(new MethodInvoker(delegate { filterByPnumber = this.tbxPNumber.Text; }));
                }




                Dictionary<string, CardHolderReportInfo> cnicDateWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();

                List<string> lstCnics = new List<string>();


                List<CCFTEvent.Event> lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                                   where
                                                       events != null && (events.EventType == 20001 || events.EventType == 20003) &&
                                                       events.OccurrenceTime >= fromDate &&
                                                       events.OccurrenceTime < toDate
                                                   select events).ToList();
                if (progressBar1.InvokeRequired)
                {
                    progress = 2;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }

                List<int> inIds = new List<int>();
                List<int> outIds = new List<int>();
                Dictionary<DateTime, double> DicTotalCheckIn = new Dictionary<DateTime, double>();
                Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlInEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();
                Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlOutEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();

                Dictionary<int, Cardholder> inCardHolders = new Dictionary<int, Cardholder>();
                Dictionary<int, Cardholder> outCardHolders = new Dictionary<int, Cardholder>();

                Dictionary<int, List<CCFTEvent.Event>> dayWiseEvents = null;

                foreach (CCFTEvent.Event events in lstEvents)
                {
                    if (events == null || events.RelatedItems == null)
                    {
                        continue;
                    }

                    foreach (RelatedItem relatedItem in events.RelatedItems)
                    {
                        if (relatedItem != null)
                        {
                            if (relatedItem.RelationCode == 0)
                            {
                                //In Events
                                if (events.EventType == 20001)
                                {
                                    inIds.Add(relatedItem.FTItemID);

                                    if (lstChlInEvents.ContainsKey(events.OccurrenceTime.Date))
                                    {
                                        if (lstChlInEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                        {
                                            if (!lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID]
                                                .Exists(ev => events.OccurrenceTime.TimeOfDay.Hours == ev.OccurrenceTime.TimeOfDay.Hours
                                                           && events.OccurrenceTime.TimeOfDay.Minutes == ev.OccurrenceTime.TimeOfDay.Minutes))
                                            {
                                                lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);
                                            }


                                        }
                                        else
                                        {

                                            lstChlInEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                        }
                                    }
                                    else
                                    {
                                        dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                        dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                        lstChlInEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                    }
                                }
                                //Out Events
                                else if (events.EventType == 20003)//Out events
                                {
                                    outIds.Add(relatedItem.FTItemID);

                                    if (lstChlOutEvents.ContainsKey(events.OccurrenceTime.Date))
                                    {
                                        if (lstChlOutEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                        {
                                            if (!lstChlOutEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Exists(ev => events.OccurrenceTime.TimeOfDay.Hours == ev.OccurrenceTime.TimeOfDay.Hours && events.OccurrenceTime.TimeOfDay.Minutes == ev.OccurrenceTime.TimeOfDay.Minutes))
                                            {
                                                lstChlOutEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);
                                            }
                                        }
                                        else
                                        {
                                            lstChlOutEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                        }
                                    }
                                    else
                                    {
                                        dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                        dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                        lstChlOutEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                    }
                                }

                            }

                        }
                    }
                }


                inCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                 where chl != null && inIds.Contains(chl.FTItemID)
                                 select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);


                List<string> strLstTempCards = (from chl in inCardHolders
                                                where chl.Value != null && (chl.Value.FirstName.ToLower().StartsWith("t-") || chl.Value.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-"))
                                                select chl.Value.LastName).ToList();



                List<CheckInAndOutInfo> filteredCheckIns = (from checkin in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                                            where checkin != null && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate &&
                                                                strLstTempCards.Contains(checkin.CardNumber) &&

                                                                //(string.IsNullOrEmpty(filterByDepartment) ||
                                                                //    ((checkin.CardHolderInfos != null &&
                                                                //    checkin.CardHolderInfos.Department != null &&
                                                                //    checkin.CardHolderInfos.Department.DepartmentName.ToLower() == filterByDepartment) ||
                                                                //    (checkin.DailyCardHolders != null &&
                                                                //    checkin.DailyCardHolders.Department.ToLower() == filterByDepartment))) &&

                                                                (string.IsNullOrEmpty(filterBySection) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Section != null &&
                                                                    checkin.CardHolderInfos.Section.SectionName.ToLower() == filterBySection) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.Section.ToLower() == filterBySection))) &&

                                                                (string.IsNullOrEmpty(filerByName) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.FirstName.ToLower().Contains(filerByName)))) &&

                                                                (string.IsNullOrEmpty(filterByCadre) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Cadre != null &&
                                                                    checkin.CardHolderInfos.Cadre.CadreName.ToLower() == filterByCadre) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.Cadre.ToLower() == filterByCadre))) &&

                                                                (string.IsNullOrEmpty(filterByCompany) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Company != null &&
                                                                    !string.IsNullOrEmpty(checkin.CardHolderInfos.Company.CompanyName) &&
                                                                    checkin.CardHolderInfos.Company.CompanyName.ToLower() == filterByCompany) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    !string.IsNullOrEmpty(checkin.DailyCardHolders.CompanyName) &&
                                                                    checkin.DailyCardHolders.CompanyName.ToLower() == filterByCompany) ||
                                                                    (checkin.Visitors != null &&
                                                                    !string.IsNullOrEmpty(checkin.Visitors.CompanyName) &&
                                                                    checkin.Visitors.CompanyName.ToLower() == filterByCompany))) &&

                                                                (string.IsNullOrEmpty(filterByCNIC) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.CNICNumber == filterByCNIC) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.CNICNumber == filterByCNIC) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.CNICNumber == filterByCNIC))) &&

                                                                (string.IsNullOrEmpty(filterByPnumber) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.PNumber == filterByPnumber)))
                                                            select checkin).ToList();



                outCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                  where chl != null && outIds.Contains(chl.FTItemID)
                                  select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);


                if (progressBar1.InvokeRequired)
                {
                    progress = 8;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }
                List<string> CadreList = new List<string>();
                CadreList.Add("MPT");
                CadreList.Add("NMPT");
                CadreList.Add("TAP");
                CadreList.Add("GTE");

                foreach (KeyValuePair<DateTime, Dictionary<int, List<CCFTEvent.Event>>> inEvent in lstChlInEvents)
                {
                    DateTime date = inEvent.Key;
                    if (inEvent.Value == null)
                    {
                        continue;
                    }

                    foreach (KeyValuePair<int, List<CCFTEvent.Event>> chlWiseEvents in inEvent.Value)
                    {
                        if (chlWiseEvents.Value == null || chlWiseEvents.Value.Count == 0 || !inCardHolders.ContainsKey(chlWiseEvents.Key))
                        {
                            continue;
                        }

                        int ftItemId = chlWiseEvents.Key;

                        Cardholder chl = inCardHolders[ftItemId];

                        if (chl == null)
                        {
                            continue;
                        }

                        bool isTempCard = chl.FirstName.ToLower().StartsWith("t-") || chl.FirstName.ToLower().StartsWith("v-") || chl.FirstName.ToLower().StartsWith("temporary-") || chl.FirstName.ToLower().StartsWith("visitor-");

                        if (isTempCard)
                        {
                            #region TempCard

                            string tempCardNumber = chl.LastName;

                            List<CheckInAndOutInfo> dateWiseCheckins = (from checkIn in filteredCheckIns
                                                                        where checkIn != null && checkIn.DateTimeIn.Date == date && checkIn.CardNumber == tempCardNumber
                                                                        select checkIn).ToList();

                            Dictionary<string, DateTime> dictInTime = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictOutTime = new Dictionary<string, DateTime>();//dateWiseCheckIn.DateTimeOut;

                            Dictionary<string, DateTime> dictCallOutInTimeAfterEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictCallOutInTimeBeforeStart = new Dictionary<string, DateTime>();

                            Dictionary<string, DateTime> dictCallOutOutTimeAfterEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictCallOutOutTimeBeforeStart = new Dictionary<string, DateTime>();


                            Dictionary<string, DateTime> dictFirstInTimeAfterDayStart = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutInTimesBeforeDayStart = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutOutTimesBeforeDayStart = new Dictionary<string, DateTime>();

                            Dictionary<string, DateTime> dictLastCallOutInTimesAfterDayEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutOutTimesAfterDayEnd = new Dictionary<string, DateTime>();

                            foreach (CheckInAndOutInfo dateWiseCheckIn in dateWiseCheckins)
                            {
                                string cnicNumber = dateWiseCheckIn.CNICNumber;
                                string firstName = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? (dateWiseCheckIn.Visitors == null ? "Unknown" : dateWiseCheckIn.Visitors.FirstName) : dateWiseCheckIn.DailyCardHolders.FirstName) : dateWiseCheckIn.CardHolderInfos.FirstName;

                                string pNumber = dateWiseCheckIn.CardHolderInfos == null || string.IsNullOrEmpty(dateWiseCheckIn.CardHolderInfos.PNumber) ? "Unknown" : dateWiseCheckIn.CardHolderInfos.PNumber;

                                string department = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Department) : (dateWiseCheckIn.CardHolderInfos.Department == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Department.DepartmentName);
                                department = string.IsNullOrEmpty(department) ? "Unknown" : department;

                                string section = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Section) : (dateWiseCheckIn.CardHolderInfos.Section == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Section.SectionName);
                                section = string.IsNullOrEmpty(section) ? "Unknown" : section;

                                string cadre = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Cadre) : (dateWiseCheckIn.CardHolderInfos.Cadre == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Cadre.CadreName);
                                cadre = string.IsNullOrEmpty(cadre) ? "Unknown" : cadre;


                                DateTime minInTime = dictInTime.ContainsKey(cnicNumber) ? dictInTime[cnicNumber] : DateTime.MaxValue;
                                DateTime maxOutTime = dictOutTime.ContainsKey(cnicNumber) ? dictOutTime[cnicNumber] : DateTime.MaxValue;

                                DateTime inDateTime = DateTime.MaxValue;
                                DateTime outDateTime = DateTime.MaxValue;

                                DateTime minCallOutInTimeAfterEnd = dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber) ? dictCallOutInTimeAfterEnd[cnicNumber] : DateTime.MaxValue;
                                DateTime minCallOutInTimeBeforeStart = dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber) ? dictCallOutInTimeBeforeStart[cnicNumber] : DateTime.MaxValue;

                                DateTime maxCallOutOutTimeAfterEnd = dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber) ? dictCallOutOutTimeAfterEnd[cnicNumber] : DateTime.MaxValue;
                                DateTime maxCallOutOutTimeBeforeStart = dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber) ? dictCallOutOutTimeBeforeStart[cnicNumber] : DateTime.MaxValue;

                                DateTime callOutInDateTime = DateTime.MaxValue;
                                DateTime callOutOutDateTime = DateTime.MaxValue;

                                DateTime firstInTimeAfterDayStart = dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber) ? dictFirstInTimeAfterDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutInTimesBeforeDayStart = dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutInTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutOutTimesBeforeDayStart = dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutOutTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;

                                DateTime lastCallOutInTimesAfterDayEnd = dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutInTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutOutTimesAfterDayEnd = dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutOutTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (dateWiseCheckIn.DateTimeIn.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = dateWiseCheckIn.DateTimeIn;

                                            if (dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutInTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                            }
                                            else
                                            {
                                                dictLastCallOutInTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                            }

                                        }

                                        callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;

                                            if (!dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                            {
                                                dictCallOutInTimeBeforeStart.Add(cnicNumber, minInTime);
                                            }

                                        }
                                        else
                                        {
                                            if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                                dictCallOutInTimeBeforeStart[cnicNumber] = minInTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < fdtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = dateWiseCheckIn.DateTimeIn;

                                                if (dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictFirstInTimeAfterDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                                }
                                                else
                                                {
                                                    dictFirstInTimeAfterDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                                }

                                            }

                                            inDateTime = dateWiseCheckIn.DateTimeIn;
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                minInTime = dateWiseCheckIn.DateTimeIn;
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                    minInTime = dateWiseCheckIn.DateTimeIn;
                                                    dictInTime[cnicNumber] = minInTime;

                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTime = dateWiseCheckIn.DateTimeIn;
                                            minCallOutInTimeAfterEnd = dateWiseCheckIn.DateTimeIn;

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < callOutInDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = callOutInDateTime;

                                                if (dictLastCallOutInTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd[cnicNumber] = callOutInDateTime;
                                                }
                                                else
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd.Add(cnicNumber, callOutInDateTime);
                                                }
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                dictCallOutInTimeAfterEnd.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minInTime;
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    if (dateWiseCheckIn.DateTimeIn.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = dateWiseCheckIn.DateTimeIn;

                                            if (dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutInTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                            }
                                            else
                                            {
                                                dictLastCallOutInTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                            }

                                        }

                                        callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                            dictCallOutInTimeBeforeStart.Add(cnicNumber, minInTime);
                                        }
                                        else
                                        {
                                            if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                                dictCallOutInTimeBeforeStart[cnicNumber] = minInTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < ndtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = dateWiseCheckIn.DateTimeIn;

                                                if (dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictFirstInTimeAfterDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                                }
                                                else
                                                {
                                                    dictFirstInTimeAfterDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                                }

                                            }

                                            inDateTime = dateWiseCheckIn.DateTimeIn;
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                minInTime = dateWiseCheckIn.DateTimeIn;
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                    minInTime = dateWiseCheckIn.DateTimeIn;
                                                    dictInTime[cnicNumber] = minInTime;

                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                            minCallOutInTimeAfterEnd = dateWiseCheckIn.DateTimeIn;

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < callOutInDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = callOutInDateTime;

                                                if (dictLastCallOutInTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd[cnicNumber] = callOutInDateTime;
                                                }
                                                else
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd.Add(cnicNumber, callOutInDateTime);
                                                }
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                dictCallOutInTimeAfterEnd.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minInTime;
                                                }
                                            }
                                        }
                                    }

                                }

                                if (minInTime == DateTime.MaxValue && minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                {
                                    continue;
                                }

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < dateWiseCheckIn.DateTimeOut)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = dateWiseCheckIn.DateTimeOut;

                                            if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeOut;
                                            }
                                            else
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeOut);
                                            }
                                        }

                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                        maxCallOutOutTimeBeforeStart = dateWiseCheckIn.DateTimeOut;

                                        if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                        {
                                            dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                        }
                                        else
                                        {
                                            dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                        }
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                                else
                                                {
                                                    callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                    maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                        if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                        }
                                                        else
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                        }
                                                    }

                                                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                    {
                                                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                    }
                                                    else
                                                    {
                                                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTime = date.Add(fdtStartDate.TimeOfDay);

                                                maxCallOutOutTimeBeforeStart = date.Add(fdtStartDate.TimeOfDay);

                                                lastCallOutOutTimesBeforeDayStart = date.Add(fdtStartDate.TimeOfDay);

                                                if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                inDateTime = date.Add(fdtStartDate.TimeOfDay);

                                                if (dictInTime.ContainsKey(cnicNumber))
                                                {
                                                    dictInTime[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictInTime.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                outDateTime = dateWiseCheckIn.DateTimeOut;
                                                maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                if (dictOutTime.ContainsKey(cnicNumber))
                                                {
                                                    dictOutTime[cnicNumber] = maxOutTime;
                                                }
                                                else
                                                {
                                                    dictOutTime.Add(cnicNumber, maxOutTime);
                                                }
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                        maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                            if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                            }
                                                            else
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                            }
                                                        }

                                                        if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                        }
                                                        else
                                                        {
                                                            dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < dateWiseCheckIn.DateTimeOut)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = dateWiseCheckIn.DateTimeOut;

                                            if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeOut;
                                            }
                                            else
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeOut);
                                            }
                                        }

                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                        maxCallOutOutTimeBeforeStart = dateWiseCheckIn.DateTimeOut;

                                        if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                        {
                                            dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                        }
                                        else
                                        {
                                            dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                        }
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                                else
                                                {
                                                    callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                    maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                        if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                        }
                                                        else
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                        }
                                                    }

                                                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                    {
                                                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                    }
                                                    else
                                                    {
                                                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTime = date.Add(ndtStartDate.TimeOfDay);

                                                maxCallOutOutTimeBeforeStart = date.Add(ndtStartDate.TimeOfDay);

                                                lastCallOutOutTimesBeforeDayStart = date.Add(ndtStartDate.TimeOfDay);

                                                if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }


                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }

                                                inDateTime = date.Add(ndtStartDate.TimeOfDay);

                                                if (dictInTime.ContainsKey(cnicNumber))
                                                {
                                                    dictInTime[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictInTime.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }

                                                outDateTime = dateWiseCheckIn.DateTimeOut;
                                                maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                if (dictOutTime.ContainsKey(cnicNumber))
                                                {
                                                    dictOutTime[cnicNumber] = maxOutTime;
                                                }
                                                else
                                                {
                                                    dictOutTime.Add(cnicNumber, maxOutTime);
                                                }
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                        maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                            if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                            }
                                                            else
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                            }
                                                        }

                                                        if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                        }
                                                        else
                                                        {
                                                            dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (maxOutTime == DateTime.MaxValue && maxCallOutOutTimeAfterEnd == DateTime.MaxValue)
                                {
                                    continue;
                                }

                                if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                {
                                    DateTime prevInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinInTime;
                                    DateTime prevOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxOutTime;

                                    DateTime prevCallOutInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinCallOutInTime;
                                    DateTime prevCallOutOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxCallOutOutTime;

                                    if (date.DayOfWeek == DayOfWeek.Friday)
                                    {

                                        if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                        {
                                            minInTime = prevInTime;

                                            if (dictInTime.ContainsKey(cnicNumber))
                                            {
                                                dictInTime[cnicNumber] = minInTime;
                                            }
                                            else
                                            {
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                        }

                                        if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            maxOutTime = prevOutTime;

                                            if (dictOutTime.ContainsKey(cnicNumber))
                                            {
                                                dictOutTime[cnicNumber] = maxOutTime;
                                            }
                                            else
                                            {
                                                dictOutTime.Add(cnicNumber, maxOutTime);
                                            }
                                        }

                                        if (prevCallOutInTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                        {
                                            if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = prevCallOutInTime;

                                                if (dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeBeforeStart[cnicNumber] = minCallOutInTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeBeforeStart.Add(cnicNumber, minCallOutInTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeAfterEnd = prevCallOutInTime;

                                                if (dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minCallOutInTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeAfterEnd.Add(cnicNumber, minCallOutInTimeAfterEnd);
                                                }
                                            }
                                        }

                                        if (prevCallOutOutTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                        {
                                            if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeBeforeStart = prevCallOutOutTime;

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (maxCallOutOutTimeAfterEnd.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeAfterEnd = prevCallOutOutTime;

                                                if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                }
                                            }
                                        }


                                        //}
                                    }
                                    else
                                    {

                                        if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                        {
                                            minInTime = prevInTime;

                                            if (dictInTime.ContainsKey(cnicNumber))
                                            {
                                                dictInTime[cnicNumber] = minInTime;
                                            }
                                            else
                                            {
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                        }

                                        if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            maxOutTime = prevOutTime;

                                            if (dictOutTime.ContainsKey(cnicNumber))
                                            {
                                                dictOutTime[cnicNumber] = maxOutTime;
                                            }
                                            else
                                            {
                                                dictOutTime.Add(cnicNumber, maxOutTime);
                                            }
                                        }

                                        if (prevCallOutInTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                        {
                                            if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = prevCallOutInTime;

                                                if (dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeBeforeStart[cnicNumber] = minCallOutInTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeBeforeStart.Add(cnicNumber, minCallOutInTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeAfterEnd = prevCallOutInTime;

                                                if (dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minCallOutInTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeAfterEnd.Add(cnicNumber, minCallOutInTimeAfterEnd);
                                                }
                                            }
                                        }

                                        if (prevCallOutOutTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                        {
                                            if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeBeforeStart = prevCallOutOutTime;

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                            {
                                                maxCallOutOutTimeAfterEnd = prevCallOutOutTime;

                                                if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                }
                                            }
                                        }


                                        //}
                                    }

                                }


                                int netNormalHours = 0;
                                int netNormalMinutes = 0;
                                int otHours = 0;
                                int otMinutes = 0;
                                int callOutHours = 0;
                                int callOutMinutes = 0;
                                string callOutFromHours = string.Empty;
                                string callOutToHours = string.Empty;
                                int lunchHours = 0;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    lunchHours = (fdtLunchEndTime - fdtLunchStartTime).Hours;

                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchStartTime)
                                    {


                                        if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours = (fdtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes = (fdtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (fdtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > fdtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours = (outDateTime.TimeOfDay - fdtLunchEndTime).Hours;
                                                    netNormalMinutes = (outDateTime.TimeOfDay - fdtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours = (fdtEndTime - fdtLunchEndTime).Hours;
                                                    netNormalMinutes = (fdtEndTime - fdtLunchEndTime).Minutes;
                                                    otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                    otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (fdtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    lunchHours = (ndtLunchEndTime - ndtLunchStartTime).Hours;

                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours = (ndtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes = (ndtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (ndtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > ndtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours = (outDateTime.TimeOfDay - ndtLunchEndTime).Hours;
                                                    netNormalMinutes = (outDateTime.TimeOfDay - ndtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours = (ndtEndTime - ndtLunchEndTime).Hours;
                                                    netNormalMinutes = (ndtEndTime - ndtLunchEndTime).Minutes;
                                                    otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                    otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (ndtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }

                                if (callOutInDateTime != DateTime.MaxValue && callOutOutDateTime != DateTime.MaxValue)
                                {
                                    callOutHours = (callOutOutDateTime - callOutInDateTime).Hours;
                                    callOutMinutes = (callOutOutDateTime - callOutInDateTime).Minutes;
                                }

                                if (minCallOutInTimeBeforeStart != DateTime.MaxValue && maxCallOutOutTimeBeforeStart != DateTime.MaxValue)
                                {
                                    callOutFromHours = minCallOutInTimeBeforeStart.ToString("HH:mm");
                                    callOutToHours = maxCallOutOutTimeBeforeStart.ToString("HH:mm");
                                }

                                if (minCallOutInTimeAfterEnd != DateTime.MaxValue && maxCallOutOutTimeAfterEnd != DateTime.MaxValue)
                                {
                                    if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                    {
                                        callOutFromHours = minCallOutInTimeAfterEnd.ToString("HH:mm");
                                    }

                                    callOutToHours = maxCallOutOutTimeAfterEnd.ToString("HH:mm");
                                }

                                if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                {
                                    CardHolderReportInfo reportInfo = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                    if (reportInfo != null)
                                    {
                                        reportInfo.NetNormalHours += netNormalHours;
                                        reportInfo.NetNormalMinutes += netNormalMinutes;
                                        reportInfo.OverTimeHours += otHours;
                                        reportInfo.OverTimeMinutes += otMinutes;
                                        reportInfo.TotalCallOutHours += callOutHours;
                                        reportInfo.TotalCallOutMinutes += callOutMinutes;

                                        if (minInTime.TimeOfDay < reportInfo.MinInTime.TimeOfDay)
                                        {
                                            reportInfo.MinInTime = minInTime;
                                        }

                                        if (maxOutTime.TimeOfDay > reportInfo.MaxOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxOutTime = maxOutTime;
                                        }

                                        if (minCallOutInTimeBeforeStart.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                        {
                                            reportInfo.MinCallOutInTime = minCallOutInTimeBeforeStart;
                                            reportInfo.CallOutFrom = callOutFromHours;
                                        }

                                        if (minCallOutInTimeAfterEnd.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                        {
                                            reportInfo.MinCallOutInTime = minCallOutInTimeAfterEnd;
                                            reportInfo.CallOutFrom = callOutFromHours;
                                        }

                                        if (maxCallOutOutTimeAfterEnd.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxCallOutOutTime = maxCallOutOutTimeAfterEnd;
                                            reportInfo.CallOutTo = callOutToHours;
                                        }

                                        if (maxCallOutOutTimeAfterEnd == DateTime.MaxValue || maxCallOutOutTimeBeforeStart.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxCallOutOutTime = maxCallOutOutTimeBeforeStart;
                                            reportInfo.CallOutTo = callOutToHours;
                                        }
                                    }
                                }
                                else
                                {
                                    lstCnics.Add(cnicNumber);

                                    cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                    {
                                        OccurrenceTime = date,
                                        FirstName = chl.FirstName,
                                        PNumber = pNumber.ToString(),
                                        CNICNumber = cnicNumber,
                                        Department = department,
                                        Section = section,
                                        Cadre = cadre,
                                        NetNormalHours = netNormalHours,
                                        OverTimeHours = otHours,
                                        TotalCallOutHours = callOutHours,
                                        NetNormalMinutes = netNormalMinutes,
                                        OverTimeMinutes = otMinutes,
                                        TotalCallOutMinutes = callOutMinutes,
                                        CallOutFrom = callOutFromHours,
                                        CallOutTo = callOutToHours,
                                        MinInTime = minInTime,
                                        MaxOutTime = maxOutTime,
                                        MinCallOutInTime = minCallOutInTimeAfterEnd < minCallOutInTimeBeforeStart ? minCallOutInTimeAfterEnd : minCallOutInTimeBeforeStart,
                                        MaxCallOutOutTime = maxCallOutOutTimeAfterEnd == DateTime.MaxValue ? maxCallOutOutTimeBeforeStart : maxCallOutOutTimeAfterEnd
                                    });


                                }
                            }

                            #endregion
                        }
                        else
                        {
                            #region Events

                            if (!lstChlOutEvents.ContainsKey(date) ||
                                lstChlOutEvents[date] == null ||
                                !lstChlOutEvents[date].ContainsKey(ftItemId) ||
                                lstChlOutEvents[date][ftItemId] == null ||
                                lstChlOutEvents[date][ftItemId].Count == 0)
                            {
                                continue;
                            }

                            List<CCFTEvent.Event> inEvents = chlWiseEvents.Value;

                            inEvents = inEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                            List<CCFTEvent.Event> outEvents = lstChlOutEvents[date][ftItemId];

                            outEvents = outEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                            int pNumber = chl.PersonalDataIntegers == null || chl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(chl.PersonalDataIntegers.ElementAt(0).Value);
                            string strPnumber = Convert.ToString(pNumber);
                            string cnicNumber = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051).Value);
                            string department = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043).Value);
                            string section = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951).Value);
                            string cadre = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952).Value);
                            string company = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059).Value);

                            strPnumber = string.IsNullOrEmpty(strPnumber) ? "Unknown" : strPnumber;
                            cnicNumber = string.IsNullOrEmpty(cnicNumber) ? "Unknown" : cnicNumber;
                            department = string.IsNullOrEmpty(department) ? "Unknown" : department;
                            section = string.IsNullOrEmpty(section) ? "Unknown" : section;
                            cadre = string.IsNullOrEmpty(cadre) ? "Unknown" : cadre;
                            company = string.IsNullOrEmpty(company) ? "Unknown" : company;

                            //if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() == filterByDepartment.ToLower())
                            //{
                            //    continue;
                            //}


                            //Filter By Section
                            if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                            {
                                continue;
                            }

                            //Filter By Cadre
                            if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && cadre.ToLower() != filterByCadre.ToLower())
                            {
                                continue;
                            }

                            //Filter By Company
                            if (!string.IsNullOrEmpty(filterByCompany) && company.ToLower() != filterByCompany.ToLower())
                            {
                                continue;
                            }

                            //Filter By CNIC
                            if (string.IsNullOrEmpty(cnicNumber) || !string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                            {
                                continue;
                            }

                            //Filter By Name
                            if (!string.IsNullOrEmpty(filerByName) && !chl.FirstName.ToLower().Contains(filerByName.ToLower()))
                            {
                                continue;
                            }

                            if (!string.IsNullOrEmpty(filterByPnumber) && strPnumber != filterByPnumber)
                            {
                                continue;
                            }

                            DateTime minInTime = DateTime.MaxValue;
                            DateTime maxOutTime = DateTime.MaxValue;

                            DateTime minCallOutInTimeAfterEnd = DateTime.MaxValue;
                            DateTime minCallOutInTimeBeforeStart = DateTime.MaxValue;

                            DateTime maxCallOutOutTimeAfterEnd = DateTime.MaxValue;
                            DateTime maxCallOutOutTimeBeforeStart = DateTime.MaxValue;

                            List<DateTime> inDateTimes = new List<DateTime>();
                            List<DateTime> outDateTimes = new List<DateTime>();

                            List<DateTime> callOutInDateTimes = new List<DateTime>();
                            List<DateTime> callOutOutDateTimes = new List<DateTime>();

                            DateTime firstInTimeAfterDayStart = DateTime.MaxValue;
                            DateTime lastCallOutInTimesBeforeDayStart = DateTime.MaxValue;
                            DateTime lastCallOutOutTimesBeforeDayStart = DateTime.MaxValue;
                            DateTime lastCallOutInTimesAfterDayEnd = DateTime.MaxValue;
                            DateTime lastCallOutOutTimesAfterDayEnd = DateTime.MaxValue;

                            foreach (CCFTEvent.Event ev in inEvents)
                            {
                                DateTime inDateTime = ev.OccurrenceTime.AddHours(5);

                                //MessageBox.Show("Event In Time: " + inDateTime.ToString());

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < inDateTime.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = inDateTime;
                                        }

                                        callOutInDateTimes.Add(inDateTime);

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = inDateTime;
                                        }
                                        else
                                        {
                                            if (inDateTime.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = inDateTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > inDateTime.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = inDateTime;
                                            }

                                            inDateTimes.Add(inDateTime);

                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                minInTime = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    minInTime = inDateTime;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTimes.Add(inDateTime);

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < inDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = inDateTime;
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                minCallOutInTimeAfterEnd = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    minCallOutInTimeAfterEnd = inDateTime;
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < inDateTime.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = inDateTime;
                                        }

                                        callOutInDateTimes.Add(inDateTime);

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = inDateTime;
                                        }
                                        else
                                        {
                                            if (inDateTime.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = inDateTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > inDateTime.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = inDateTime;
                                            }

                                            inDateTimes.Add(inDateTime);
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                minInTime = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                    minInTime = inDateTime;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTimes.Add(inDateTime);

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < inDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = inDateTime;
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                minCallOutInTimeAfterEnd = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    minCallOutInTimeAfterEnd = inDateTime;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (minInTime == DateTime.MaxValue && minCallOutInTimeAfterEnd == DateTime.MaxValue)
                            {
                                continue;
                            }

                            foreach (CCFTEvent.Event ev in outEvents)
                            {
                                DateTime outDateTime = ev.OccurrenceTime.AddHours(5);

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < outDateTime)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = outDateTime;
                                        }

                                        callOutOutDateTimes.Add(outDateTime);

                                        maxCallOutOutTimeBeforeStart = outDateTime;
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                outDateTimes.Add(outDateTime);
                                                if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                {
                                                    maxOutTime = outDateTime;
                                                }

                                            }
                                            else
                                            {
                                                if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    maxOutTime = outDateTime;
                                                }
                                                else
                                                {
                                                    callOutOutDateTimes.Add(outDateTime);

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                    }

                                                    maxCallOutOutTimeAfterEnd = outDateTime;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));
                                                maxCallOutOutTimeBeforeStart = date.Add(fdtStartDate.TimeOfDay);
                                                lastCallOutOutTimesBeforeDayStart = date.Add(fdtStartDate.TimeOfDay);

                                                inDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));
                                                minInTime = date.Add(fdtStartDate.TimeOfDay);

                                                outDateTimes.Add(outDateTime);
                                                maxOutTime = outDateTime;
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                    {
                                                        maxOutTime = outDateTime;
                                                    }

                                                }
                                                else
                                                {
                                                    if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTimes.Add(outDateTime);
                                                        maxOutTime = outDateTime;
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTimes.Add(outDateTime);

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                        }

                                                        maxCallOutOutTimeAfterEnd = outDateTime;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < outDateTime)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = outDateTime;
                                        }

                                        callOutOutDateTimes.Add(outDateTime);
                                        maxCallOutOutTimeBeforeStart = outDateTime;
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                outDateTimes.Add(outDateTime);
                                                if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                {
                                                    maxOutTime = outDateTime;
                                                }

                                            }
                                            else
                                            {
                                                if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    maxOutTime = outDateTime;
                                                }
                                                else
                                                {
                                                    callOutOutDateTimes.Add(outDateTime);

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                    }

                                                    maxCallOutOutTimeAfterEnd = outDateTime;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));
                                                maxCallOutOutTimeBeforeStart = date.Add(ndtStartDate.TimeOfDay);
                                                lastCallOutOutTimesBeforeDayStart = date.Add(ndtStartDate.TimeOfDay);

                                                inDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));
                                                minInTime = date.Add(ndtStartDate.TimeOfDay);

                                                outDateTimes.Add(outDateTime);
                                                maxOutTime = outDateTime;
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    outDateTimes.Add(outDateTime);

                                                    if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                    {
                                                        maxOutTime = outDateTime;
                                                    }
                                                }
                                                else
                                                {
                                                    if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTimes.Add(outDateTime);
                                                        maxOutTime = outDateTime;
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTimes.Add(outDateTime);

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                        }

                                                        maxCallOutOutTimeAfterEnd = outDateTime;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (maxOutTime == DateTime.MaxValue && maxCallOutOutTimeAfterEnd == DateTime.MaxValue)
                            {
                                continue;
                            }

                            if (lastCallOutInTimesAfterDayEnd != DateTime.MaxValue)
                            {
                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd > lastCallOutOutTimesAfterDayEnd)
                                {
                                    CCFTEvent.Event missingOutEvent = (from events in lstEvents
                                                                       where events != null &&
                                                                             events.EventType == 20003 &&
                                                                             events.RelatedItems != null &&
                                                                                 (from relatedItem in events.RelatedItems
                                                                                  where relatedItem != null &&
                                                                                        relatedItem.RelationCode == 0 &&
                                                                                        relatedItem.FTItemID == ftItemId
                                                                                  select relatedItem).Any() &&
                                                                             events.OccurrenceTime.Date == date.AddDays(1)
                                                                       select events).FirstOrDefault();


                                    if (missingOutEvent != null)
                                    {
                                        DateTime outDateTime = missingOutEvent.OccurrenceTime.AddHours(5);

                                        if (date.AddDays(1).DayOfWeek == DayOfWeek.Friday)
                                        {
                                            if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                            {
                                                callOutOutDateTimes.Add(outDateTime);

                                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                {
                                                    lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                }

                                                maxCallOutOutTimeAfterEnd = outDateTime;
                                            }
                                            else
                                            {
                                                callOutOutDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));

                                                lastCallOutOutTimesAfterDayEnd = date.Add(fdtStartDate.TimeOfDay);

                                                maxCallOutOutTimeAfterEnd = date.Add(fdtStartDate.TimeOfDay);
                                            }
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                            {
                                                callOutOutDateTimes.Add(outDateTime);

                                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                {
                                                    lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                }

                                                maxCallOutOutTimeAfterEnd = outDateTime;
                                            }
                                            else
                                            {
                                                callOutOutDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));

                                                lastCallOutOutTimesAfterDayEnd = date.Add(ndtStartDate.TimeOfDay);

                                                maxCallOutOutTimeAfterEnd = date.Add(ndtStartDate.TimeOfDay);
                                            }
                                        }
                                    }
                                }
                            }


                            if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                DateTime prevInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinInTime;
                                DateTime prevOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxOutTime;

                                DateTime prevCallOutInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinCallOutInTime;
                                DateTime prevCallOutOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxCallOutOutTime;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
                                        minInTime = prevInTime;
                                    }

                                    if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        maxOutTime = prevOutTime;
                                    }
                                    //}
                                    //else
                                    //{
                                    if (prevCallOutInTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }
                                    else
                                    {
                                        if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }

                                    if (prevCallOutOutTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            maxCallOutOutTimeBeforeStart = prevCallOutOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                        {
                                            maxCallOutOutTimeAfterEnd = prevCallOutOutTime;
                                        }
                                    }


                                    //}
                                }
                                else
                                {
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
                                        minInTime = prevInTime;
                                    }

                                    if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        maxOutTime = prevOutTime;
                                    }
                                    //}
                                    //else
                                    //{
                                    if (prevCallOutInTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }
                                    else
                                    {
                                        if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }

                                    if (prevCallOutOutTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            maxCallOutOutTimeBeforeStart = prevCallOutOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                        {
                                            maxCallOutOutTimeAfterEnd = prevCallOutOutTime;
                                        }
                                    }


                                    //}
                                }

                            }

                            int netNormalHours = 0;
                            int netNormalMinutes = 0;
                            int otHours = 0;
                            int otMinutes = 0;
                            int callOutHours = 0;
                            int callOutMinutes = 0;
                            string callOutFromHours = string.Empty;
                            string callOutToHours = string.Empty;
                            int lunchHours = 0;

                            inDateTimes.OrderBy((a) => a.TimeOfDay);
                            outDateTimes.OrderBy((a) => a.TimeOfDay);

                            foreach (DateTime inDateTime in inDateTimes)
                            {
                                //MessageBox.Show(this, "In Time: " + inDateTime.ToString());
                                DateTime outDateTime = DateTime.MaxValue;

                                //finding nearest out time wrt in time.
                                foreach (DateTime oDateTime in outDateTimes)
                                {
                                    if (oDateTime.TimeOfDay < inDateTime.TimeOfDay)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (oDateTime.TimeOfDay < outDateTime.TimeOfDay)
                                        {
                                            outDateTime = oDateTime;
                                        }
                                    }
                                }

                                //MessageBox.Show(this, "Out Time: " + outDateTime.ToString());

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    lunchHours = (fdtLunchEndTime - fdtLunchStartTime).Hours;
                                    //MessageBox.Show("Lunch Hours: " + lunchHours);
                                    //MessageBox.Show("In Hours: " + inTime.ToString());
                                    //MessageBox.Show("Out Hours: " + outTime.ToString());

                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours += (fdtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes += (fdtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours += (fdtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > fdtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours += (outDateTime.TimeOfDay - fdtLunchEndTime).Hours;
                                                    netNormalMinutes += (outDateTime.TimeOfDay - fdtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours += (fdtEndTime - fdtLunchEndTime).Hours;
                                                    netNormalMinutes += (fdtEndTime - fdtLunchEndTime).Minutes;
                                                    otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                    otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours += (fdtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    lunchHours = (ndtLunchEndTime - ndtLunchStartTime).Hours;

                                    //MessageBox.Show(this, "Lunch Hrs: " + lunchHours);

                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours += (ndtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes += (ndtLunchStartTime - inDateTime.TimeOfDay).Minutes;

                                            //MessageBox.Show(this, "ibl obl Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;

                                                //MessageBox.Show(this, "ibl oal obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }
                                            else
                                            {
                                                netNormalHours += (ndtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                //MessageBox.Show(this, "ibl oal oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > ndtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= ndtWithBeforeGraceTimeEndTime)
                                                {
                                                    netNormalHours += (outDateTime.TimeOfDay - ndtLunchEndTime).Hours;
                                                    netNormalMinutes += (outDateTime.TimeOfDay - ndtLunchEndTime).Minutes;

                                                    //MessageBox.Show(this, "ible oale obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                                }
                                                else
                                                {
                                                    netNormalHours += (ndtEndTime - ndtLunchEndTime).Hours;
                                                    netNormalMinutes += (ndtEndTime - ndtLunchEndTime).Minutes;
                                                    otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                    otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                    //MessageBox.Show(this, "ible oale oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;

                                                //MessageBox.Show(this, "iale obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }
                                            else
                                            {
                                                netNormalHours += (ndtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                //MessageBox.Show(this, "iale oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }

                                        }
                                    }
                                }
                            }

                            callOutInDateTimes.OrderBy((a) => a);
                            callOutOutDateTimes.OrderBy((a) => a);

                            foreach (DateTime callOutInDateTime in callOutInDateTimes)
                            {
                                DateTime callOutOutDateTime = DateTime.MaxValue;

                                //finding nearest out time wrt in time.
                                foreach (DateTime oDateTime in callOutOutDateTimes)
                                {
                                    if (oDateTime < callOutInDateTime)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (oDateTime < callOutOutDateTime)
                                        {
                                            callOutOutDateTime = oDateTime;
                                        }
                                    }
                                }

                                if (callOutInDateTime != DateTime.MaxValue && callOutOutDateTime != DateTime.MaxValue)
                                {
                                    callOutHours += (callOutOutDateTime - callOutInDateTime).Hours;
                                    callOutMinutes += (callOutOutDateTime - callOutInDateTime).Minutes;
                                }
                            }


                            if (minCallOutInTimeBeforeStart != DateTime.MaxValue && maxCallOutOutTimeBeforeStart != DateTime.MaxValue)
                            {
                                callOutFromHours = minCallOutInTimeBeforeStart.ToString("HH:mm");
                                callOutToHours = maxCallOutOutTimeBeforeStart.ToString("HH:mm");
                            }

                            if (minCallOutInTimeAfterEnd != DateTime.MaxValue && maxCallOutOutTimeAfterEnd != DateTime.MaxValue)
                            {
                                if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                {
                                    callOutFromHours = minCallOutInTimeAfterEnd.ToString("HH:mm");
                                }

                                callOutToHours = maxCallOutOutTimeAfterEnd.ToString("HH:mm");
                            }

                            if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                CardHolderReportInfo reportInfo = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                if (reportInfo != null)
                                {
                                    reportInfo.NetNormalHours += netNormalHours;
                                    reportInfo.NetNormalMinutes += netNormalMinutes;
                                    reportInfo.OverTimeHours += otHours;
                                    reportInfo.OverTimeMinutes += otMinutes;
                                    reportInfo.TotalCallOutHours += callOutHours;
                                    reportInfo.TotalCallOutMinutes += callOutMinutes;

                                    if (minInTime.TimeOfDay < reportInfo.MinInTime.TimeOfDay)
                                    {
                                        reportInfo.MinInTime = minInTime;
                                    }

                                    if (maxOutTime.TimeOfDay > reportInfo.MaxOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxOutTime = maxOutTime;
                                    }

                                    if (minCallOutInTimeBeforeStart.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                    {
                                        reportInfo.MinCallOutInTime = minCallOutInTimeBeforeStart;
                                        reportInfo.CallOutFrom = callOutFromHours;
                                    }

                                    if (minCallOutInTimeAfterEnd.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                    {
                                        reportInfo.MinCallOutInTime = minCallOutInTimeAfterEnd;
                                        reportInfo.CallOutFrom = callOutFromHours;
                                    }

                                    if (maxCallOutOutTimeAfterEnd.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxCallOutOutTime = maxCallOutOutTimeAfterEnd;
                                        reportInfo.CallOutTo = callOutToHours;
                                    }

                                    if (maxCallOutOutTimeAfterEnd == DateTime.MaxValue || maxCallOutOutTimeBeforeStart.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxCallOutOutTime = maxCallOutOutTimeBeforeStart;
                                        reportInfo.CallOutTo = callOutToHours;
                                    }
                                }
                            }
                            else
                            {
                                lstCnics.Add(cnicNumber);

                                cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                {
                                    OccurrenceTime = date,
                                    FirstName = chl.FirstName,
                                    PNumber = pNumber.ToString(),
                                    CNICNumber = cnicNumber,
                                    Department = department,
                                    Section = section,
                                    Cadre = cadre,
                                    NetNormalHours = netNormalHours,
                                    OverTimeHours = otHours,
                                    TotalCallOutHours = callOutHours,
                                    NetNormalMinutes = netNormalMinutes,
                                    OverTimeMinutes = otMinutes,
                                    TotalCallOutMinutes = callOutMinutes,
                                    CallOutFrom = callOutFromHours,
                                    CallOutTo = callOutToHours,
                                    MinInTime = minInTime,
                                    MaxOutTime = maxOutTime,
                                    MinCallOutInTime = minCallOutInTimeAfterEnd < minCallOutInTimeBeforeStart ? minCallOutInTimeAfterEnd : minCallOutInTimeBeforeStart,
                                    MaxCallOutOutTime = maxCallOutOutTimeAfterEnd == DateTime.MaxValue ? maxCallOutOutTimeBeforeStart : maxCallOutOutTimeAfterEnd
                                });





                            }

                            #endregion
                        }
                    }
                }

                if (progressBar1.InvokeRequired)
                {
                    progress = 12;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }
                #endregion
                Dictionary<string, Dictionary<string, double>> CadreNic = new Dictionary<string, Dictionary<string, double>>();
                Dictionary<string, Dictionary<string, Dictionary<string, double>>> Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                Depart_Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();


                if (cnicDateWiseReportInfo != null && cnicDateWiseReportInfo.Keys.Count > 0)
                {
                    int totalDays = (toDate.Date - fromDate.Date).Days;

                    for (int i = 0; i <= totalDays; i++)
                    {
                        DateTime date = fromDate.Date.AddDays(i);

                        foreach (string strCnic in lstCnics)
                        {
                            if (cnicDateWiseReportInfo.ContainsKey(strCnic + "^" + date.ToString()))
                            {
                                continue;
                            }
                            else
                            {
                                CardHolderReportInfo reportInfo = (from cnicDate in cnicDateWiseReportInfo
                                                                   where cnicDate.Key.Contains(strCnic)
                                                                   select cnicDate.Value).FirstOrDefault();

                                if (reportInfo != null)
                                {
                                    cnicDateWiseReportInfo.Add(strCnic + "^" + date.ToString(), new CardHolderReportInfo()
                                    {
                                        OccurrenceTime = date,
                                        FirstName = reportInfo.FirstName,
                                        PNumber = reportInfo.PNumber,
                                        CNICNumber = reportInfo.CNICNumber,
                                        Department = reportInfo.Department,
                                        Section = reportInfo.Section,
                                        Cadre = reportInfo.Cadre,
                                        MinInTime = DateTime.MaxValue,
                                        MaxOutTime = DateTime.MaxValue,
                                        MinCallOutInTime = DateTime.MaxValue,
                                        MaxCallOutOutTime = DateTime.MaxValue
                                    });


                                }
                            }
                        }
                    }


                    List<Cardholder> remainingCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                                             where chl != null &&
                                                                  !(from pds in chl.PersonalDataStrings
                                                                    where pds != null && pds.PersonalDataFieldID == 5051 && pds.Value != null && lstCnics.Contains(pds.Value)
                                                                    select pds).Any()
                                                             select chl).ToList();

                    foreach (Cardholder remainingChl in remainingCardHolders)
                    {
                        int pNumber = remainingChl.PersonalDataIntegers == null || remainingChl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(remainingChl.PersonalDataIntegers.ElementAt(0).Value);
                        string strPnumber = Convert.ToString(pNumber);
                        string cnicNumber = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051).Value);
                        string department = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043).Value);
                        string section = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951).Value);
                        string cadre = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952).Value);
                        string company = remainingChl.PersonalDataStrings == null ? "Unknown" : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059) == null ? "Unknown" : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059).Value);

                        //if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() == filterByDepartment.ToLower())
                        //{
                        //    continue;
                        //}


                        //Filter By Section
                        if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                        {
                            continue;
                        }

                        //Filter By Cadre
                        if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && cadre.ToLower() != filterByCadre.ToLower())
                        {
                            continue;
                        }

                        //Filter By Company
                        if (!string.IsNullOrEmpty(filterByCompany) && company.ToLower() != filterByCompany.ToLower())
                        {
                            continue;
                        }

                        //Filter By CNIC
                        if (string.IsNullOrEmpty(cnicNumber) || !string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                        {
                            continue;
                        }

                        //Filter By Name
                        if (!string.IsNullOrEmpty(filerByName) && !remainingChl.FirstName.ToLower().Contains(filerByName.ToLower()))
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(filterByPnumber) && strPnumber != filterByPnumber)
                        {
                            continue;
                        }

                        for (int i = 0; i <= totalDays; i++)
                        {
                            DateTime date = fromDate.Date.AddDays(i);
                            if (!cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                {
                                    OccurrenceTime = date,
                                    FirstName = remainingChl.FirstName,
                                    PNumber = strPnumber,
                                    CNICNumber = cnicNumber,
                                    Department = department,
                                    Section = section,
                                    Cadre = cadre,
                                    MinInTime = DateTime.MaxValue,
                                    MaxOutTime = DateTime.MaxValue,
                                    MinCallOutInTime = DateTime.MaxValue,
                                    MaxCallOutOutTime = DateTime.MaxValue
                                });




                            }

                        }
                    }

                    this.Summary_Report = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                    Dictionary<string, double> CnicHoursEffert = new Dictionary<string, double>();
                    //Dictionary<string, double> CnicHoursOthers = new Dictionary<string, double>();
                    int diff1 = 0;

                    int otherWorkers = 0;
                   
                    Depart_Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                    foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicDateWiseReportInfo)
                    {

                        if (progress <= 80)
                        {
                            if (progressBar1.InvokeRequired)
                            {
                                progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress++; }));
                            }
                        }

                        if (reportInfo.Value == null)
                        {
                            continue;
                        }

                        string cnicNumber = reportInfo.Value.CNICNumber;
                        string department = reportInfo.Value.Department;
                        string section = reportInfo.Value.Section;
                        string cadre = reportInfo.Value.Cadre;



                        if (reportInfo.Value.Department == null)
                        {
                            continue;
                        }



                        #region Making Hours

                        department = department.ToUpper();


                        if (reportInfo.Value.MinInTime != DateTime.MaxValue)
                        {
                            if (reportInfo.Value.MaxOutTime != DateTime.MaxValue)
                            {
                                diff1 = Convert.ToInt32((reportInfo.Value.MaxOutTime - reportInfo.Value.MinInTime).TotalHours);


                                if (CadreList.Contains(cadre))
                                {
                                    if (Summary_Report.ContainsKey(department))
                                    {
                                        if (Summary_Report[department].ContainsKey("EFFERT"))
                                        {
                                            if (Summary_Report[department]["EFFERT"].ContainsKey("HoursCount"))
                                            {
                                                Summary_Report[department]["EFFERT"]["HoursCount"] = Summary_Report[department]["EFFERT"]["HoursCount"] + diff1;
                                                Summary_Report[department]["EFFERT"]["NoOfEmployee"] = Summary_Report[department]["EFFERT"]["NoOfEmployee"] + 1;

                                            }
                                            else
                                            {
                                                Summary_Report[department]["EFFERT"].Add("HoursCount", diff1);
                                                otherWorkers++;
                                                Summary_Report[department]["EFFERT"].Add("NoOfEmployee", otherWorkers);
                                            }
                                        }
                                        else
                                        {
                                            CnicHoursEffert = new Dictionary<string, double>();
                                            CnicHoursEffert.Add("HoursCount", diff1);
                                            otherWorkers = 0;
                                            otherWorkers++;
                                            CnicHoursEffert.Add("NoOfEmployee", otherWorkers);

                                            Summary_Report[department].Add("EFFERT", CnicHoursEffert);

                                        }
                                    }
                                    else
                                    {
                                        CnicHoursEffert = new Dictionary<string, double>();
                                        CnicHoursEffert.Add("HoursCount", diff1);
                                        otherWorkers = 0;
                                        otherWorkers++;
                                        CnicHoursEffert.Add("NoOfEmployee", otherWorkers);


                                        CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                        CadreNic.Add("EFFERT", CnicHoursEffert);

                                        Summary_Report.Add(department, CadreNic);

                                    }



                                }
                                else
                                {
                                    //OTHERS
                                    if (Summary_Report.ContainsKey(department))
                                    {
                                        if (Summary_Report[department].ContainsKey("OTHERS"))
                                        {
                                            if (Summary_Report[department]["OTHERS"].ContainsKey("HoursCount"))
                                            {
                                                Summary_Report[department]["OTHERS"]["HoursCount"] = Summary_Report[department]["OTHERS"]["HoursCount"] + diff1;
                                                Summary_Report[department]["OTHERS"]["NoOfEmployee"] = Summary_Report[department]["OTHERS"]["NoOfEmployee"] + 1;

                                            }
                                            else
                                            {
                                                Summary_Report[department]["OTHERS"].Add("HoursCount", diff1);
                                                otherWorkers++;
                                                Summary_Report[department]["OTHERS"].Add("NoOfEmployee", otherWorkers);
                                            }
                                        }
                                        else
                                        {
                                            CnicHoursEffert = new Dictionary<string, double>();
                                            CnicHoursEffert.Add("HoursCount", diff1);
                                            otherWorkers = 0;
                                            otherWorkers++;
                                            CnicHoursEffert.Add("NoOfEmployee", otherWorkers);

                                            Summary_Report[department].Add("OTHERS", CnicHoursEffert);

                                        }
                                    }
                                    else
                                    {
                                        CnicHoursEffert = new Dictionary<string, double>();
                                        CnicHoursEffert.Add("HoursCount", diff1);
                                        otherWorkers = 0;
                                        otherWorkers++;
                                        CnicHoursEffert.Add("NoOfEmployee", otherWorkers);


                                        CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                        CadreNic.Add("OTHERS", CnicHoursEffert);


                                        Summary_Report.Add(department, CadreNic);

                                    }




                                }
                            }
                        }

                        #endregion

                       
                       

                    }
                }



                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = false; }));
                }
            }
            catch (Exception ex)
            {
                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = false; }));
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = 0; }));
                }
                string aa = ex.StackTrace.ToString();
                CreateLogFiles.ErrorLog(ex.Message + " " + ex.StackTrace);
                MessageBox.Show(ex.Message);
            }
        }

        public void Generate_ManHours_Detail_report()
        {
            try
            {
               
                int progress = 0;

                #region Actual Data
                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = true; }));
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }


                this.Depart_Date_CadreNic = null;
                DateTime fromDate = this.dtpFromDate.Value.Date;
                DateTime fromDateUtc = fromDate.ToUniversalTime();
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                DateTime toDateUtc = toDate.ToUniversalTime();

                DateTime ndtStartDate = this.dtpNdtStart.Value;
                DateTime ndtEndDate = this.dtpNdtEnd.Value;
                DateTime ndtLunchStartDate = this.dtpNdtLunchStart.Value;
                DateTime ndtLunchEndDate = this.dtpNdtLunchEnd.Value;

                DateTime fdtStartDate = this.dtpFdtStart.Value;
                DateTime fdtEndDate = this.dtpFdtEnd.Value;
                DateTime fdtLunchStartDate = this.dtpFdtLunchStart.Value;
                DateTime fdtLunchEndDate = this.dtpFdtLunchEnd.Value;

                int ndtGraceTimeBeforeStart = Convert.ToInt32(nudNdtGraceTimeBeforeStart.Value);
                int ndtGraceTimeAfterStart = Convert.ToInt32(nudNdtGraceTimeBeforeStart.Value);
                int ndtGraceTimeBeforeEnd = Convert.ToInt32(nudNdtGraceTimeBeforeEnd.Value);
                int ndtGraceTimeAfterEnd = Convert.ToInt32(nudNdtGraceTimeBeforeEnd.Value);
                int ndtGraceTimeBeforeLunchStart = Convert.ToInt32(nudNdtGraceTimeBeforeLunchStart.Value);
                int ndtGraceTimeAfterLunchStart = Convert.ToInt32(nudNdtGraceTimeBeforeLunchStart.Value);
                int ndtGraceTimeBeforeLunchEnd = Convert.ToInt32(nudNdtGraceTimeBeforeLunchEnd.Value);
                int ndtGraceTimeAfterLunchEnd = Convert.ToInt32(nudNdtGraceTimeBeforeLunchEnd.Value);

                int fdtGraceTimeBeforeStart = Convert.ToInt32(nudFdtGraceTimeBeforeStart.Value);
                int fdtGraceTimeAfterStart = Convert.ToInt32(nudFdtGraceTimeBeforeStart.Value);
                int fdtGraceTimeBeforeEnd = Convert.ToInt32(nudFdtGraceTimeBeforeEnd.Value);
                int fdtGraceTimeAfterEnd = Convert.ToInt32(nudFdtGraceTimeBeforeEnd.Value);
                int fdtGraceTimeBeforeLunchStart = Convert.ToInt32(nudFdtGraceTimeBeforeLunchStart.Value);
                int fdtGraceTimeAfterLunchStart = Convert.ToInt32(nudFdtGraceTimeBeforeLunchStart.Value);
                int fdtGraceTimeBeforeLunchEnd = Convert.ToInt32(nudFdtGraceTimeBeforeLunchEnd.Value);
                int fdtGraceTimeAfterLunchEnd = Convert.ToInt32(nudFdtGraceTimeBeforeLunchEnd.Value);

                TimeSpan ndtStartTime = this.dtpNdtStart.Value.TimeOfDay;
                TimeSpan ndtEndTime = this.dtpNdtEnd.Value.TimeOfDay;
                TimeSpan ndtLunchStartTime = this.dtpNdtLunchStart.Value.TimeOfDay;
                TimeSpan ndtLunchEndTime = this.dtpNdtLunchEnd.Value.TimeOfDay;

                TimeSpan fdtStartTime = this.dtpFdtStart.Value.TimeOfDay;
                TimeSpan fdtEndTime = this.dtpFdtEnd.Value.TimeOfDay;
                TimeSpan fdtLunchStartTime = this.dtpFdtLunchStart.Value.TimeOfDay;
                TimeSpan fdtLunchEndTime = this.dtpFdtLunchEnd.Value.TimeOfDay;

                TimeSpan ndtWithBeforeGraceTimeStartTime = this.dtpNdtStart.Value.AddMinutes(ndtGraceTimeBeforeStart * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeEndTime = this.dtpNdtEnd.Value.AddMinutes(ndtGraceTimeBeforeEnd * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeLunchStartTime = this.dtpNdtLunchStart.Value.AddMinutes(ndtGraceTimeBeforeLunchStart * -1).TimeOfDay;
                TimeSpan ndtWithBeforeGraceTimeLunchEndTime = this.dtpNdtLunchEnd.Value.AddMinutes(ndtGraceTimeBeforeLunchEnd * -1).TimeOfDay;

                TimeSpan ndtWithAfterGraceTimeStartTime = this.dtpNdtStart.Value.AddMinutes(ndtGraceTimeAfterStart).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeEndTime = this.dtpNdtEnd.Value.AddMinutes(ndtGraceTimeAfterEnd).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeLunchStartTime = this.dtpNdtLunchStart.Value.AddMinutes(ndtGraceTimeAfterLunchStart).TimeOfDay;
                TimeSpan ndtWithAfterGraceTimeLunchEndTime = this.dtpNdtLunchEnd.Value.AddMinutes(ndtGraceTimeAfterLunchEnd).TimeOfDay;

                TimeSpan fdtWithBeforeGraceTimeStartTime = this.dtpFdtStart.Value.AddMinutes(fdtGraceTimeBeforeStart * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeEndTime = this.dtpFdtEnd.Value.AddMinutes(fdtGraceTimeBeforeEnd * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeLunchStartTime = this.dtpFdtLunchStart.Value.AddMinutes(fdtGraceTimeBeforeLunchStart * -1).TimeOfDay;
                TimeSpan fdtWithBeforeGraceTimeLunchEndTime = this.dtpFdtLunchEnd.Value.AddMinutes(fdtGraceTimeBeforeLunchEnd * -1).TimeOfDay;

                TimeSpan fdtWithAfterGraceTimeStartTime = this.dtpFdtStart.Value.AddMinutes(fdtGraceTimeAfterStart).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeEndTime = this.dtpFdtEnd.Value.AddMinutes(fdtGraceTimeAfterEnd).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeLunchStartTime = this.dtpFdtLunchStart.Value.AddMinutes(fdtGraceTimeAfterLunchStart).TimeOfDay;
                TimeSpan fdtWithAfterGraceTimeLunchEndTime = this.dtpFdtLunchEnd.Value.AddMinutes(fdtGraceTimeAfterLunchEnd).TimeOfDay;

                string filterByDepartment = "";
                string filterBySection = "";
                string filerByName = "";
                string filterByCadre = "";
                string filterByCompany = "";
                string filterByCNIC = "";
                string filterByPnumber = "";

                if (cbxDepartments.InvokeRequired)
                {
                    cbxDepartments.Invoke(new MethodInvoker(delegate { filterByDepartment = this.cbxDepartments.Text.ToLower(); }));
                }

                if (filterByDepartment.ToUpper() == "EFERTDHKALL")
                {
                    MessageBox.Show("This Report is Not for All Department.Select any Specific Department.");

                }

                if (cbxSections.InvokeRequired)
                {
                    cbxSections.Invoke(new MethodInvoker(delegate { filterBySection = this.cbxSections.Text.ToLower(); }));
                }


                if (tbxName.InvokeRequired)
                {
                    tbxName.Invoke(new MethodInvoker(delegate { filerByName = this.tbxName.Text.ToLower(); }));
                }


                if (cbxCadre.InvokeRequired)
                {
                    cbxCadre.Invoke(new MethodInvoker(delegate { filterByCadre = this.cbxCadre.Text.ToLower(); }));
                }


                if (cbxCompany.InvokeRequired)
                {
                    cbxCompany.Invoke(new MethodInvoker(delegate { filterByCompany = this.cbxCompany.Text.ToLower(); }));
                }


                if (tbxCnic.InvokeRequired)
                {
                    tbxCnic.Invoke(new MethodInvoker(delegate { filterByCNIC = this.tbxCnic.Text; }));
                }


                if (tbxPNumber.InvokeRequired)
                {
                    tbxPNumber.Invoke(new MethodInvoker(delegate { filterByPnumber = this.tbxPNumber.Text; }));
                }




                Dictionary<string, CardHolderReportInfo> cnicDateWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();

                List<string> lstCnics = new List<string>();






                List<CCFTEvent.Event> lstEvents = new List<CCFTEvent.Event>();
                try
                {
                     lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                                       where
                                                           events != null && (events.EventType == 20001 || events.EventType == 20003) &&
                                                           events.OccurrenceTime >= fromDate &&
                                                           events.OccurrenceTime < toDate
                                                       select events).ToList();
                }
                catch(Exception ex)
                {
                    CreateLogFiles.ErrorLog(ex.Message + " " + ex.StackTrace);
                    Generate_ManHours_Detail_report();
                }
                if (progressBar1.InvokeRequired)
                {
                    progress = 2;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }

                List<int> inIds = new List<int>();
                List<int> outIds = new List<int>();
                Dictionary<DateTime, double> DicTotalCheckIn = new Dictionary<DateTime, double>();
                Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlInEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();
                Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlOutEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();

                Dictionary<int, Cardholder> inCardHolders = new Dictionary<int, Cardholder>();
                Dictionary<int, Cardholder> outCardHolders = new Dictionary<int, Cardholder>();

                Dictionary<int, List<CCFTEvent.Event>> dayWiseEvents = null;

                foreach (CCFTEvent.Event events in lstEvents)
                {
                    if (events == null || events.RelatedItems == null)
                    {
                        continue;
                    }

                    foreach (RelatedItem relatedItem in events.RelatedItems)
                    {
                        if (relatedItem != null)
                        {
                            if (relatedItem.RelationCode == 0)
                            {
                                //In Events
                                if (events.EventType == 20001)
                                {
                                    inIds.Add(relatedItem.FTItemID);

                                    if (lstChlInEvents.ContainsKey(events.OccurrenceTime.Date))
                                    {
                                        if (lstChlInEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                        {
                                            if (!lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID]
                                                .Exists(ev => events.OccurrenceTime.TimeOfDay.Hours == ev.OccurrenceTime.TimeOfDay.Hours
                                                           && events.OccurrenceTime.TimeOfDay.Minutes == ev.OccurrenceTime.TimeOfDay.Minutes))
                                            {
                                                lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);
                                            }


                                        }
                                        else
                                        {

                                            lstChlInEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                        }
                                    }
                                    else
                                    {
                                        dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                        dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                        lstChlInEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                    }
                                }
                                //Out Events
                                else if (events.EventType == 20003)//Out events
                                {
                                    outIds.Add(relatedItem.FTItemID);

                                    if (lstChlOutEvents.ContainsKey(events.OccurrenceTime.Date))
                                    {
                                        if (lstChlOutEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                        {
                                            if (!lstChlOutEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Exists(ev => events.OccurrenceTime.TimeOfDay.Hours == ev.OccurrenceTime.TimeOfDay.Hours && events.OccurrenceTime.TimeOfDay.Minutes == ev.OccurrenceTime.TimeOfDay.Minutes))
                                            {
                                                lstChlOutEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);
                                            }
                                        }
                                        else
                                        {
                                            lstChlOutEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                        }
                                    }
                                    else
                                    {
                                        dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                        dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                        lstChlOutEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                    }
                                }

                            }

                        }
                    }
                }


                inCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                 where chl != null && inIds.Contains(chl.FTItemID)
                                 select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);


                List<string> strLstTempCards = (from chl in inCardHolders
                                                where chl.Value != null && (chl.Value.FirstName.ToLower().StartsWith("t-") || chl.Value.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-"))
                                                select chl.Value.LastName).ToList();



                List<CheckInAndOutInfo> filteredCheckIns = (from checkin in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                                            where checkin != null && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate &&
                                                                strLstTempCards.Contains(checkin.CardNumber) &&

                                                                //(string.IsNullOrEmpty(filterByDepartment) ||
                                                                //    ((checkin.CardHolderInfos != null &&
                                                                //    checkin.CardHolderInfos.Department != null &&
                                                                //    checkin.CardHolderInfos.Department.DepartmentName.ToLower() == filterByDepartment) ||
                                                                //    (checkin.DailyCardHolders != null &&
                                                                //    checkin.DailyCardHolders.Department.ToLower() == filterByDepartment))) &&

                                                                (string.IsNullOrEmpty(filterBySection) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Section != null &&
                                                                    checkin.CardHolderInfos.Section.SectionName.ToLower() == filterBySection) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.Section.ToLower() == filterBySection))) &&

                                                                (string.IsNullOrEmpty(filerByName) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.FirstName.ToLower().Contains(filerByName)))) &&

                                                                (string.IsNullOrEmpty(filterByCadre) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Cadre != null &&
                                                                    checkin.CardHolderInfos.Cadre.CadreName.ToLower() == filterByCadre) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.Cadre.ToLower() == filterByCadre))) &&

                                                                (string.IsNullOrEmpty(filterByCompany) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Company != null &&
                                                                    !string.IsNullOrEmpty(checkin.CardHolderInfos.Company.CompanyName) &&
                                                                    checkin.CardHolderInfos.Company.CompanyName.ToLower() == filterByCompany) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    !string.IsNullOrEmpty(checkin.DailyCardHolders.CompanyName) &&
                                                                    checkin.DailyCardHolders.CompanyName.ToLower() == filterByCompany) ||
                                                                    (checkin.Visitors != null &&
                                                                    !string.IsNullOrEmpty(checkin.Visitors.CompanyName) &&
                                                                    checkin.Visitors.CompanyName.ToLower() == filterByCompany))) &&

                                                                (string.IsNullOrEmpty(filterByCNIC) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.CNICNumber == filterByCNIC) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.CNICNumber == filterByCNIC) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.CNICNumber == filterByCNIC))) &&

                                                                (string.IsNullOrEmpty(filterByPnumber) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.PNumber == filterByPnumber)))
                                                            select checkin).ToList();



                outCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                  where chl != null && outIds.Contains(chl.FTItemID)
                                  select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);


                if (progressBar1.InvokeRequired)
                {
                    progress = 8;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }
                List<string> CadreList = new List<string>();
                CadreList.Add("MPT");
                CadreList.Add("NMPT");
                CadreList.Add("TAP");
                CadreList.Add("GTE");

                foreach (KeyValuePair<DateTime, Dictionary<int, List<CCFTEvent.Event>>> inEvent in lstChlInEvents)
                {
                    DateTime date = inEvent.Key;
                    if (inEvent.Value == null)
                    {
                        continue;
                    }

                    foreach (KeyValuePair<int, List<CCFTEvent.Event>> chlWiseEvents in inEvent.Value)
                    {
                        if (chlWiseEvents.Value == null || chlWiseEvents.Value.Count == 0 || !inCardHolders.ContainsKey(chlWiseEvents.Key))
                        {
                            continue;
                        }

                        int ftItemId = chlWiseEvents.Key;

                        Cardholder chl = inCardHolders[ftItemId];

                        if (chl == null)
                        {
                            continue;
                        }

                        bool isTempCard = chl.FirstName.ToLower().StartsWith("t-") || chl.FirstName.ToLower().StartsWith("v-") || chl.FirstName.ToLower().StartsWith("temporary-") || chl.FirstName.ToLower().StartsWith("visitor-");

                        if (isTempCard)
                        {
                            #region TempCard

                            string tempCardNumber = chl.LastName;

                            List<CheckInAndOutInfo> dateWiseCheckins = (from checkIn in filteredCheckIns
                                                                        where checkIn != null && checkIn.DateTimeIn.Date == date && checkIn.CardNumber == tempCardNumber
                                                                        select checkIn).ToList();

                            Dictionary<string, DateTime> dictInTime = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictOutTime = new Dictionary<string, DateTime>();//dateWiseCheckIn.DateTimeOut;

                            Dictionary<string, DateTime> dictCallOutInTimeAfterEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictCallOutInTimeBeforeStart = new Dictionary<string, DateTime>();

                            Dictionary<string, DateTime> dictCallOutOutTimeAfterEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictCallOutOutTimeBeforeStart = new Dictionary<string, DateTime>();


                            Dictionary<string, DateTime> dictFirstInTimeAfterDayStart = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutInTimesBeforeDayStart = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutOutTimesBeforeDayStart = new Dictionary<string, DateTime>();

                            Dictionary<string, DateTime> dictLastCallOutInTimesAfterDayEnd = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictLastCallOutOutTimesAfterDayEnd = new Dictionary<string, DateTime>();

                            foreach (CheckInAndOutInfo dateWiseCheckIn in dateWiseCheckins)
                            {
                                string cnicNumber = dateWiseCheckIn.CNICNumber;
                                string firstName = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? (dateWiseCheckIn.Visitors == null ? "Unknown" : dateWiseCheckIn.Visitors.FirstName) : dateWiseCheckIn.DailyCardHolders.FirstName) : dateWiseCheckIn.CardHolderInfos.FirstName;

                                string pNumber = dateWiseCheckIn.CardHolderInfos == null || string.IsNullOrEmpty(dateWiseCheckIn.CardHolderInfos.PNumber) ? "Unknown" : dateWiseCheckIn.CardHolderInfos.PNumber;

                                string department = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Department) : (dateWiseCheckIn.CardHolderInfos.Department == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Department.DepartmentName);
                                department = string.IsNullOrEmpty(department) ? "Unknown" : department;

                                string section = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Section) : (dateWiseCheckIn.CardHolderInfos.Section == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Section.SectionName);
                                section = string.IsNullOrEmpty(section) ? "Unknown" : section;

                                string cadre = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? "Unknown" : dateWiseCheckIn.DailyCardHolders.Cadre) : (dateWiseCheckIn.CardHolderInfos.Cadre == null ? "Unknown" : dateWiseCheckIn.CardHolderInfos.Cadre.CadreName);
                                cadre = string.IsNullOrEmpty(cadre) ? "Unknown" : cadre;


                                DateTime minInTime = dictInTime.ContainsKey(cnicNumber) ? dictInTime[cnicNumber] : DateTime.MaxValue;
                                DateTime maxOutTime = dictOutTime.ContainsKey(cnicNumber) ? dictOutTime[cnicNumber] : DateTime.MaxValue;

                                DateTime inDateTime = DateTime.MaxValue;
                                DateTime outDateTime = DateTime.MaxValue;

                                DateTime minCallOutInTimeAfterEnd = dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber) ? dictCallOutInTimeAfterEnd[cnicNumber] : DateTime.MaxValue;
                                DateTime minCallOutInTimeBeforeStart = dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber) ? dictCallOutInTimeBeforeStart[cnicNumber] : DateTime.MaxValue;

                                DateTime maxCallOutOutTimeAfterEnd = dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber) ? dictCallOutOutTimeAfterEnd[cnicNumber] : DateTime.MaxValue;
                                DateTime maxCallOutOutTimeBeforeStart = dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber) ? dictCallOutOutTimeBeforeStart[cnicNumber] : DateTime.MaxValue;

                                DateTime callOutInDateTime = DateTime.MaxValue;
                                DateTime callOutOutDateTime = DateTime.MaxValue;

                                DateTime firstInTimeAfterDayStart = dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber) ? dictFirstInTimeAfterDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutInTimesBeforeDayStart = dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutInTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutOutTimesBeforeDayStart = dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutOutTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;

                                DateTime lastCallOutInTimesAfterDayEnd = dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutInTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;
                                DateTime lastCallOutOutTimesAfterDayEnd = dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber) ? dictLastCallOutOutTimesBeforeDayStart[cnicNumber] : DateTime.MaxValue;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (dateWiseCheckIn.DateTimeIn.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = dateWiseCheckIn.DateTimeIn;

                                            if (dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutInTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                            }
                                            else
                                            {
                                                dictLastCallOutInTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                            }

                                        }

                                        callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;

                                            if (!dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                            {
                                                dictCallOutInTimeBeforeStart.Add(cnicNumber, minInTime);
                                            }
                                        }
                                        else
                                        {
                                            if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                                dictCallOutInTimeBeforeStart[cnicNumber] = minInTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < fdtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = dateWiseCheckIn.DateTimeIn;

                                                if (dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictFirstInTimeAfterDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                                }
                                                else
                                                {
                                                    dictFirstInTimeAfterDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                                }

                                            }

                                            inDateTime = dateWiseCheckIn.DateTimeIn;
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                minInTime = dateWiseCheckIn.DateTimeIn;
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                    minInTime = dateWiseCheckIn.DateTimeIn;
                                                    dictInTime[cnicNumber] = minInTime;

                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTime = dateWiseCheckIn.DateTimeIn;
                                            minCallOutInTimeAfterEnd = dateWiseCheckIn.DateTimeIn;

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < callOutInDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = callOutInDateTime;

                                                if (dictLastCallOutInTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd[cnicNumber] = callOutInDateTime;
                                                }
                                                else
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd.Add(cnicNumber, callOutInDateTime);
                                                }
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                dictCallOutInTimeAfterEnd.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minInTime;
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    if (dateWiseCheckIn.DateTimeIn.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = dateWiseCheckIn.DateTimeIn;

                                            if (dictLastCallOutInTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutInTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                            }
                                            else
                                            {
                                                dictLastCallOutInTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                            }

                                        }

                                        callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                            dictCallOutInTimeBeforeStart.Add(cnicNumber, minInTime);
                                        }
                                        else
                                        {
                                            if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = dateWiseCheckIn.DateTimeIn;
                                                dictCallOutInTimeBeforeStart[cnicNumber] = minInTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < ndtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > dateWiseCheckIn.DateTimeIn.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = dateWiseCheckIn.DateTimeIn;

                                                if (dictFirstInTimeAfterDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictFirstInTimeAfterDayStart[cnicNumber] = dateWiseCheckIn.DateTimeIn;
                                                }
                                                else
                                                {
                                                    dictFirstInTimeAfterDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeIn);
                                                }

                                            }

                                            inDateTime = dateWiseCheckIn.DateTimeIn;
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                minInTime = dateWiseCheckIn.DateTimeIn;
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set T: " + inTime.ToString());
                                                    minInTime = dateWiseCheckIn.DateTimeIn;
                                                    dictInTime[cnicNumber] = minInTime;

                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTime = dateWiseCheckIn.DateTimeIn;

                                            minCallOutInTimeAfterEnd = dateWiseCheckIn.DateTimeIn;

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < callOutInDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = callOutInDateTime;

                                                if (dictLastCallOutInTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd[cnicNumber] = callOutInDateTime;
                                                }
                                                else
                                                {
                                                    dictLastCallOutInTimesAfterDayEnd.Add(cnicNumber, callOutInDateTime);
                                                }
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                dictCallOutInTimeAfterEnd.Add(cnicNumber, minInTime);
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minInTime;
                                                }
                                            }
                                        }
                                    }

                                }

                                if (minInTime == DateTime.MaxValue && minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                {
                                    continue;
                                }

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < dateWiseCheckIn.DateTimeOut)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = dateWiseCheckIn.DateTimeOut;

                                            if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeOut;
                                            }
                                            else
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeOut);
                                            }
                                        }

                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                        maxCallOutOutTimeBeforeStart = dateWiseCheckIn.DateTimeOut;

                                        if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                        {
                                            dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                        }
                                        else
                                        {
                                            dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                        }
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                                else
                                                {
                                                    callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                    maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                        if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                        }
                                                        else
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                        }
                                                    }

                                                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                    {
                                                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                    }
                                                    else
                                                    {
                                                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTime = date.Add(fdtStartDate.TimeOfDay);

                                                maxCallOutOutTimeBeforeStart = date.Add(fdtStartDate.TimeOfDay);

                                                lastCallOutOutTimesBeforeDayStart = date.Add(fdtStartDate.TimeOfDay);

                                                if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                inDateTime = date.Add(fdtStartDate.TimeOfDay);

                                                if (dictInTime.ContainsKey(cnicNumber))
                                                {
                                                    dictInTime[cnicNumber] = date.Add(fdtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictInTime.Add(cnicNumber, date.Add(fdtStartDate.TimeOfDay));
                                                }

                                                outDateTime = dateWiseCheckIn.DateTimeOut;
                                                maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                if (dictOutTime.ContainsKey(cnicNumber))
                                                {
                                                    dictOutTime[cnicNumber] = maxOutTime;
                                                }
                                                else
                                                {
                                                    dictOutTime.Add(cnicNumber, maxOutTime);
                                                }
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                        maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                            if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                            }
                                                            else
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                            }
                                                        }

                                                        if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                        }
                                                        else
                                                        {
                                                            dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < dateWiseCheckIn.DateTimeOut)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = dateWiseCheckIn.DateTimeOut;

                                            if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = dateWiseCheckIn.DateTimeOut;
                                            }
                                            else
                                            {
                                                dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, dateWiseCheckIn.DateTimeOut);
                                            }
                                        }

                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                        maxCallOutOutTimeBeforeStart = dateWiseCheckIn.DateTimeOut;

                                        if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                        {
                                            dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                        }
                                        else
                                        {
                                            dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                        }
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTime = dateWiseCheckIn.DateTimeOut;
                                                    maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                    if (dictOutTime.ContainsKey(cnicNumber))
                                                    {
                                                        dictOutTime[cnicNumber] = maxOutTime;
                                                    }
                                                    else
                                                    {
                                                        dictOutTime.Add(cnicNumber, maxOutTime);
                                                    }
                                                }
                                                else
                                                {
                                                    callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                    maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                        if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                        }
                                                        else
                                                        {
                                                            dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                        }
                                                    }

                                                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                    {
                                                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                    }
                                                    else
                                                    {
                                                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTime = date.Add(ndtStartDate.TimeOfDay);

                                                maxCallOutOutTimeBeforeStart = date.Add(ndtStartDate.TimeOfDay);

                                                lastCallOutOutTimesBeforeDayStart = date.Add(ndtStartDate.TimeOfDay);

                                                if (dictLastCallOutOutTimesBeforeDayStart.ContainsKey(cnicNumber))
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictLastCallOutOutTimesBeforeDayStart.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }


                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }

                                                inDateTime = date.Add(ndtStartDate.TimeOfDay);

                                                if (dictInTime.ContainsKey(cnicNumber))
                                                {
                                                    dictInTime[cnicNumber] = date.Add(ndtStartDate.TimeOfDay);
                                                }
                                                else
                                                {
                                                    dictInTime.Add(cnicNumber, date.Add(ndtStartDate.TimeOfDay));
                                                }

                                                outDateTime = dateWiseCheckIn.DateTimeOut;
                                                maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                if (dictOutTime.ContainsKey(cnicNumber))
                                                {
                                                    dictOutTime[cnicNumber] = maxOutTime;
                                                }
                                                else
                                                {
                                                    dictOutTime.Add(cnicNumber, maxOutTime);
                                                }
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay > minInTime.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (dateWiseCheckIn.DateTimeOut.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTime = dateWiseCheckIn.DateTimeOut;
                                                        maxOutTime = dateWiseCheckIn.DateTimeOut;

                                                        if (dictOutTime.ContainsKey(cnicNumber))
                                                        {
                                                            dictOutTime[cnicNumber] = maxOutTime;
                                                        }
                                                        else
                                                        {
                                                            dictOutTime.Add(cnicNumber, maxOutTime);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTime = dateWiseCheckIn.DateTimeOut;

                                                        maxCallOutOutTimeAfterEnd = dateWiseCheckIn.DateTimeOut;

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < callOutOutDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                                                            if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                                                            }
                                                            else
                                                            {
                                                                dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                                                            }
                                                        }

                                                        if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                        {
                                                            dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                        }
                                                        else
                                                        {
                                                            dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (maxOutTime == DateTime.MaxValue && maxCallOutOutTimeAfterEnd == DateTime.MaxValue)
                                {
                                    continue;
                                }

                                if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                {
                                    DateTime prevInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinInTime;
                                    DateTime prevOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxOutTime;

                                    DateTime prevCallOutInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinCallOutInTime;
                                    DateTime prevCallOutOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxCallOutOutTime;

                                    if (date.DayOfWeek == DayOfWeek.Friday)
                                    {

                                        if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                        {
                                            minInTime = prevInTime;

                                            if (dictInTime.ContainsKey(cnicNumber))
                                            {
                                                dictInTime[cnicNumber] = minInTime;
                                            }
                                            else
                                            {
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                        }

                                        if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            maxOutTime = prevOutTime;

                                            if (dictOutTime.ContainsKey(cnicNumber))
                                            {
                                                dictOutTime[cnicNumber] = maxOutTime;
                                            }
                                            else
                                            {
                                                dictOutTime.Add(cnicNumber, maxOutTime);
                                            }
                                        }

                                        if (prevCallOutInTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                        {
                                            if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = prevCallOutInTime;

                                                if (dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeBeforeStart[cnicNumber] = minCallOutInTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeBeforeStart.Add(cnicNumber, minCallOutInTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeAfterEnd = prevCallOutInTime;

                                                if (dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minCallOutInTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeAfterEnd.Add(cnicNumber, minCallOutInTimeAfterEnd);
                                                }
                                            }
                                        }

                                        if (prevCallOutOutTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                        {
                                            if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeBeforeStart = prevCallOutOutTime;

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (maxCallOutOutTimeAfterEnd.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeAfterEnd = prevCallOutOutTime;

                                                if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                }
                                            }
                                        }


                                        //}
                                    }
                                    else
                                    {

                                        if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                        {
                                            minInTime = prevInTime;

                                            if (dictInTime.ContainsKey(cnicNumber))
                                            {
                                                dictInTime[cnicNumber] = minInTime;
                                            }
                                            else
                                            {
                                                dictInTime.Add(cnicNumber, minInTime);
                                            }
                                        }

                                        if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            maxOutTime = prevOutTime;

                                            if (dictOutTime.ContainsKey(cnicNumber))
                                            {
                                                dictOutTime[cnicNumber] = maxOutTime;
                                            }
                                            else
                                            {
                                                dictOutTime.Add(cnicNumber, maxOutTime);
                                            }
                                        }

                                        if (prevCallOutInTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                        {
                                            if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = prevCallOutInTime;

                                                if (dictCallOutInTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeBeforeStart[cnicNumber] = minCallOutInTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeBeforeStart.Add(cnicNumber, minCallOutInTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                            {
                                                minCallOutInTimeAfterEnd = prevCallOutInTime;

                                                if (dictCallOutInTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutInTimeAfterEnd[cnicNumber] = minCallOutInTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutInTimeAfterEnd.Add(cnicNumber, minCallOutInTimeAfterEnd);
                                                }
                                            }
                                        }

                                        if (prevCallOutOutTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                        {
                                            if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                            {
                                                maxCallOutOutTimeBeforeStart = prevCallOutOutTime;

                                                if (dictCallOutOutTimeBeforeStart.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeBeforeStart[cnicNumber] = maxCallOutOutTimeBeforeStart;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeBeforeStart.Add(cnicNumber, maxCallOutOutTimeBeforeStart);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                            {
                                                maxCallOutOutTimeAfterEnd = prevCallOutOutTime;

                                                if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                                                {
                                                    dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                                                }
                                                else
                                                {
                                                    dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                                                }
                                            }
                                        }


                                        //}
                                    }

                                }


                                int netNormalHours = 0;
                                int netNormalMinutes = 0;
                                int otHours = 0;
                                int otMinutes = 0;
                                int callOutHours = 0;
                                int callOutMinutes = 0;
                                string callOutFromHours = string.Empty;
                                string callOutToHours = string.Empty;
                                int lunchHours = 0;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    lunchHours = (fdtLunchEndTime - fdtLunchStartTime).Hours;

                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchStartTime)
                                    {


                                        if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours = (fdtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes = (fdtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (fdtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > fdtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours = (outDateTime.TimeOfDay - fdtLunchEndTime).Hours;
                                                    netNormalMinutes = (outDateTime.TimeOfDay - fdtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours = (fdtEndTime - fdtLunchEndTime).Hours;
                                                    netNormalMinutes = (fdtEndTime - fdtLunchEndTime).Minutes;
                                                    otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                    otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (fdtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    lunchHours = (ndtLunchEndTime - ndtLunchStartTime).Hours;

                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours = (ndtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes = (ndtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (ndtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes = (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > ndtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours = (outDateTime.TimeOfDay - ndtLunchEndTime).Hours;
                                                    netNormalMinutes = (outDateTime.TimeOfDay - ndtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours = (ndtEndTime - ndtLunchEndTime).Hours;
                                                    netNormalMinutes = (ndtEndTime - ndtLunchEndTime).Minutes;
                                                    otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                    otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours = (ndtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes = (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours = (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes = (outDateTime.TimeOfDay - ndtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }

                                if (callOutInDateTime != DateTime.MaxValue && callOutOutDateTime != DateTime.MaxValue)
                                {
                                    callOutHours = (callOutOutDateTime - callOutInDateTime).Hours;
                                    callOutMinutes = (callOutOutDateTime - callOutInDateTime).Minutes;
                                }

                                if (minCallOutInTimeBeforeStart != DateTime.MaxValue && maxCallOutOutTimeBeforeStart != DateTime.MaxValue)
                                {
                                    callOutFromHours = minCallOutInTimeBeforeStart.ToString("HH:mm");
                                    callOutToHours = maxCallOutOutTimeBeforeStart.ToString("HH:mm");
                                }

                                if (minCallOutInTimeAfterEnd != DateTime.MaxValue && maxCallOutOutTimeAfterEnd != DateTime.MaxValue)
                                {
                                    if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                    {
                                        callOutFromHours = minCallOutInTimeAfterEnd.ToString("HH:mm");
                                    }

                                    callOutToHours = maxCallOutOutTimeAfterEnd.ToString("HH:mm");
                                }

                                if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                {
                                    CardHolderReportInfo reportInfo = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                    if (reportInfo != null)
                                    {
                                        reportInfo.NetNormalHours += netNormalHours;
                                        reportInfo.NetNormalMinutes += netNormalMinutes;
                                        reportInfo.OverTimeHours += otHours;
                                        reportInfo.OverTimeMinutes += otMinutes;
                                        reportInfo.TotalCallOutHours += callOutHours;
                                        reportInfo.TotalCallOutMinutes += callOutMinutes;

                                        if (minInTime.TimeOfDay < reportInfo.MinInTime.TimeOfDay)
                                        {
                                            reportInfo.MinInTime = minInTime;
                                        }

                                        if (maxOutTime.TimeOfDay > reportInfo.MaxOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxOutTime = maxOutTime;
                                        }

                                        if (minCallOutInTimeBeforeStart.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                        {
                                            reportInfo.MinCallOutInTime = minCallOutInTimeBeforeStart;
                                            reportInfo.CallOutFrom = callOutFromHours;
                                        }

                                        if (minCallOutInTimeAfterEnd.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                        {
                                            reportInfo.MinCallOutInTime = minCallOutInTimeAfterEnd;
                                            reportInfo.CallOutFrom = callOutFromHours;
                                        }

                                        if (maxCallOutOutTimeAfterEnd.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxCallOutOutTime = maxCallOutOutTimeAfterEnd;
                                            reportInfo.CallOutTo = callOutToHours;
                                        }

                                        if (maxCallOutOutTimeAfterEnd == DateTime.MaxValue || maxCallOutOutTimeBeforeStart.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                        {
                                            reportInfo.MaxCallOutOutTime = maxCallOutOutTimeBeforeStart;
                                            reportInfo.CallOutTo = callOutToHours;
                                        }
                                    }
                                }
                                else
                                {
                                    lstCnics.Add(cnicNumber);

                                    cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                    {
                                        OccurrenceTime = date,
                                        FirstName = chl.FirstName,
                                        PNumber = pNumber.ToString(),
                                        CNICNumber = cnicNumber,
                                        Department = department,
                                        Section = section,
                                        Cadre = cadre,
                                        NetNormalHours = netNormalHours,
                                        OverTimeHours = otHours,
                                        TotalCallOutHours = callOutHours,
                                        NetNormalMinutes = netNormalMinutes,
                                        OverTimeMinutes = otMinutes,
                                        TotalCallOutMinutes = callOutMinutes,
                                        CallOutFrom = callOutFromHours,
                                        CallOutTo = callOutToHours,
                                        MinInTime = minInTime,
                                        MaxOutTime = maxOutTime,
                                        MinCallOutInTime = minCallOutInTimeAfterEnd < minCallOutInTimeBeforeStart ? minCallOutInTimeAfterEnd : minCallOutInTimeBeforeStart,
                                        MaxCallOutOutTime = maxCallOutOutTimeAfterEnd == DateTime.MaxValue ? maxCallOutOutTimeBeforeStart : maxCallOutOutTimeAfterEnd
                                    });


                                }
                            }

                            #endregion
                        }
                        else
                        {
                            #region Events

                            if (!lstChlOutEvents.ContainsKey(date) ||
                                lstChlOutEvents[date] == null ||
                                !lstChlOutEvents[date].ContainsKey(ftItemId) ||
                                lstChlOutEvents[date][ftItemId] == null ||
                                lstChlOutEvents[date][ftItemId].Count == 0)
                            {
                                continue;
                            }

                            List<CCFTEvent.Event> inEvents = chlWiseEvents.Value;

                            inEvents = inEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                            List<CCFTEvent.Event> outEvents = lstChlOutEvents[date][ftItemId];

                            outEvents = outEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                            int pNumber = chl.PersonalDataIntegers == null || chl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(chl.PersonalDataIntegers.ElementAt(0).Value);
                            string strPnumber = Convert.ToString(pNumber);
                            string cnicNumber = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051).Value);
                            string department = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043).Value);
                            string section = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951).Value);
                            string cadre = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952).Value);
                            string company = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059).Value);

                            strPnumber = string.IsNullOrEmpty(strPnumber) ? "Unknown" : strPnumber;
                            cnicNumber = string.IsNullOrEmpty(cnicNumber) ? "Unknown" : cnicNumber;
                            department = string.IsNullOrEmpty(department) ? "Unknown" : department;
                            section = string.IsNullOrEmpty(section) ? "Unknown" : section;
                            cadre = string.IsNullOrEmpty(cadre) ? "Unknown" : cadre;
                            company = string.IsNullOrEmpty(company) ? "Unknown" : company;

                            //Filter By Department
                            //if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() != filterByDepartment.ToLower())
                            //{
                            //    continue;
                            //}

                            //Filter By Section
                            if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                            {
                                continue;
                            }

                            //Filter By Cadre
                            if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && cadre.ToLower() != filterByCadre.ToLower())
                            {
                                continue;
                            }

                            //Filter By Company
                            if (!string.IsNullOrEmpty(filterByCompany) && company.ToLower() != filterByCompany.ToLower())
                            {
                                continue;
                            }

                            //Filter By CNIC
                            if (string.IsNullOrEmpty(cnicNumber) || !string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                            {
                                continue;
                            }

                            //Filter By Name
                            if (!string.IsNullOrEmpty(filerByName) && !chl.FirstName.ToLower().Contains(filerByName.ToLower()))
                            {
                                continue;
                            }

                            if (!string.IsNullOrEmpty(filterByPnumber) && strPnumber != filterByPnumber)
                            {
                                continue;
                            }

                            DateTime minInTime = DateTime.MaxValue;
                            DateTime maxOutTime = DateTime.MaxValue;

                            DateTime minCallOutInTimeAfterEnd = DateTime.MaxValue;
                            DateTime minCallOutInTimeBeforeStart = DateTime.MaxValue;

                            DateTime maxCallOutOutTimeAfterEnd = DateTime.MaxValue;
                            DateTime maxCallOutOutTimeBeforeStart = DateTime.MaxValue;

                            List<DateTime> inDateTimes = new List<DateTime>();
                            List<DateTime> outDateTimes = new List<DateTime>();

                            List<DateTime> callOutInDateTimes = new List<DateTime>();
                            List<DateTime> callOutOutDateTimes = new List<DateTime>();

                            DateTime firstInTimeAfterDayStart = DateTime.MaxValue;
                            DateTime lastCallOutInTimesBeforeDayStart = DateTime.MaxValue;
                            DateTime lastCallOutOutTimesBeforeDayStart = DateTime.MaxValue;
                            DateTime lastCallOutInTimesAfterDayEnd = DateTime.MaxValue;
                            DateTime lastCallOutOutTimesAfterDayEnd = DateTime.MaxValue;

                            foreach (CCFTEvent.Event ev in inEvents)
                            {
                                DateTime inDateTime = ev.OccurrenceTime.AddHours(5);

                                //MessageBox.Show("Event In Time: " + inDateTime.ToString());

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < inDateTime.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = inDateTime;
                                        }

                                        callOutInDateTimes.Add(inDateTime);

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = inDateTime;
                                        }
                                        else
                                        {
                                            if (inDateTime.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = inDateTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > inDateTime.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = inDateTime;
                                            }

                                            inDateTimes.Add(inDateTime);

                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                minInTime = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    minInTime = inDateTime;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTimes.Add(inDateTime);

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < inDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = inDateTime;
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                minCallOutInTimeAfterEnd = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    minCallOutInTimeAfterEnd = inDateTime;
                                                }
                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue || lastCallOutInTimesBeforeDayStart.TimeOfDay < inDateTime.TimeOfDay)
                                        {
                                            lastCallOutInTimesBeforeDayStart = inDateTime;
                                        }

                                        callOutInDateTimes.Add(inDateTime);

                                        if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                        {
                                            minCallOutInTimeBeforeStart = inDateTime;
                                        }
                                        else
                                        {
                                            if (inDateTime.TimeOfDay < minCallOutInTimeBeforeStart.TimeOfDay)
                                            {
                                                minCallOutInTimeBeforeStart = inDateTime;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeEndTime)
                                        {
                                            if (firstInTimeAfterDayStart == DateTime.MaxValue || firstInTimeAfterDayStart.TimeOfDay > inDateTime.TimeOfDay)
                                            {
                                                firstInTimeAfterDayStart = inDateTime;
                                            }

                                            inDateTimes.Add(inDateTime);
                                            if (minInTime == DateTime.MaxValue)
                                            {
                                                //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                minInTime = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minInTime.TimeOfDay)
                                                {
                                                    //MessageBox.Show("In Hours set: " + inTime.ToString());
                                                    minInTime = inDateTime;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            callOutInDateTimes.Add(inDateTime);

                                            if (lastCallOutInTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd < inDateTime)
                                            {
                                                lastCallOutInTimesAfterDayEnd = inDateTime;
                                            }

                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                minCallOutInTimeAfterEnd = inDateTime;
                                            }
                                            else
                                            {
                                                if (inDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    minCallOutInTimeAfterEnd = inDateTime;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (minInTime == DateTime.MaxValue && minCallOutInTimeAfterEnd == DateTime.MaxValue)
                            {
                                continue;
                            }

                            foreach (CCFTEvent.Event ev in outEvents)
                            {
                                DateTime outDateTime = ev.OccurrenceTime.AddHours(5);

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < outDateTime)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = outDateTime;
                                        }

                                        callOutOutDateTimes.Add(outDateTime);

                                        maxCallOutOutTimeBeforeStart = outDateTime;
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                outDateTimes.Add(outDateTime);
                                                if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                {
                                                    maxOutTime = outDateTime;
                                                }

                                            }
                                            else
                                            {
                                                if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    maxOutTime = outDateTime;
                                                }
                                                else
                                                {
                                                    callOutOutDateTimes.Add(outDateTime);

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                    }

                                                    maxCallOutOutTimeAfterEnd = outDateTime;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));
                                                maxCallOutOutTimeBeforeStart = date.Add(fdtStartDate.TimeOfDay);
                                                lastCallOutOutTimesBeforeDayStart = date.Add(fdtStartDate.TimeOfDay);

                                                inDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));
                                                minInTime = date.Add(fdtStartDate.TimeOfDay);

                                                outDateTimes.Add(outDateTime);
                                                maxOutTime = outDateTime;
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                    {
                                                        maxOutTime = outDateTime;
                                                    }

                                                }
                                                else
                                                {
                                                    if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTimes.Add(outDateTime);
                                                        maxOutTime = outDateTime;
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTimes.Add(outDateTime);

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                        }

                                                        maxCallOutOutTimeAfterEnd = outDateTime;
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (lastCallOutOutTimesBeforeDayStart == DateTime.MaxValue || lastCallOutOutTimesBeforeDayStart < outDateTime)
                                        {
                                            lastCallOutOutTimesBeforeDayStart = outDateTime;
                                        }

                                        callOutOutDateTimes.Add(outDateTime);
                                        maxCallOutOutTimeBeforeStart = outDateTime;
                                    }
                                    else
                                    {
                                        if (lastCallOutInTimesBeforeDayStart == DateTime.MaxValue)
                                        {
                                            if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                            {
                                                outDateTimes.Add(outDateTime);
                                                if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                {
                                                    maxOutTime = outDateTime;
                                                }

                                            }
                                            else
                                            {
                                                if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                {
                                                    outDateTimes.Add(outDateTime);
                                                    maxOutTime = outDateTime;
                                                }
                                                else
                                                {
                                                    callOutOutDateTimes.Add(outDateTime);

                                                    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                    {
                                                        lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                    }

                                                    maxCallOutOutTimeAfterEnd = outDateTime;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (lastCallOutInTimesBeforeDayStart.TimeOfDay > lastCallOutOutTimesBeforeDayStart.TimeOfDay)
                                            {
                                                callOutOutDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));
                                                maxCallOutOutTimeBeforeStart = date.Add(ndtStartDate.TimeOfDay);
                                                lastCallOutOutTimesBeforeDayStart = date.Add(ndtStartDate.TimeOfDay);

                                                inDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));
                                                minInTime = date.Add(ndtStartDate.TimeOfDay);

                                                outDateTimes.Add(outDateTime);
                                                maxOutTime = outDateTime;
                                            }
                                            else
                                            {
                                                if (minCallOutInTimeAfterEnd == DateTime.MaxValue)
                                                {
                                                    outDateTimes.Add(outDateTime);

                                                    if (maxOutTime == DateTime.MaxValue || outDateTime.TimeOfDay > maxOutTime.TimeOfDay)
                                                    {
                                                        maxOutTime = outDateTime;
                                                    }
                                                }
                                                else
                                                {
                                                    if (outDateTime.TimeOfDay < minCallOutInTimeAfterEnd.TimeOfDay)
                                                    {
                                                        outDateTimes.Add(outDateTime);
                                                        maxOutTime = outDateTime;
                                                    }
                                                    else
                                                    {
                                                        callOutOutDateTimes.Add(outDateTime);

                                                        if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                        {
                                                            lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                        }

                                                        maxCallOutOutTimeAfterEnd = outDateTime;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (maxOutTime == DateTime.MaxValue && maxCallOutOutTimeAfterEnd == DateTime.MaxValue)
                            {
                                continue;
                            }

                            if (lastCallOutInTimesAfterDayEnd != DateTime.MaxValue)
                            {
                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd > lastCallOutOutTimesAfterDayEnd)
                                {
                                    CCFTEvent.Event missingOutEvent = (from events in lstEvents
                                                                       where events != null &&
                                                                             events.EventType == 20003 &&
                                                                             events.RelatedItems != null &&
                                                                                 (from relatedItem in events.RelatedItems
                                                                                  where relatedItem != null &&
                                                                                        relatedItem.RelationCode == 0 &&
                                                                                        relatedItem.FTItemID == ftItemId
                                                                                  select relatedItem).Any() &&
                                                                             events.OccurrenceTime.Date == date.AddDays(1)
                                                                       select events).FirstOrDefault();


                                    if (missingOutEvent != null)
                                    {
                                        DateTime outDateTime = missingOutEvent.OccurrenceTime.AddHours(5);

                                        if (date.AddDays(1).DayOfWeek == DayOfWeek.Friday)
                                        {
                                            if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                            {
                                                callOutOutDateTimes.Add(outDateTime);

                                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                {
                                                    lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                }

                                                maxCallOutOutTimeAfterEnd = outDateTime;
                                            }
                                            else
                                            {
                                                callOutOutDateTimes.Add(date.Add(fdtStartDate.TimeOfDay));

                                                lastCallOutOutTimesAfterDayEnd = date.Add(fdtStartDate.TimeOfDay);

                                                maxCallOutOutTimeAfterEnd = date.Add(fdtStartDate.TimeOfDay);
                                            }
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                            {
                                                callOutOutDateTimes.Add(outDateTime);

                                                if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutOutTimesAfterDayEnd < outDateTime)
                                                {
                                                    lastCallOutOutTimesAfterDayEnd = outDateTime;
                                                }

                                                maxCallOutOutTimeAfterEnd = outDateTime;
                                            }
                                            else
                                            {
                                                callOutOutDateTimes.Add(date.Add(ndtStartDate.TimeOfDay));

                                                lastCallOutOutTimesAfterDayEnd = date.Add(ndtStartDate.TimeOfDay);

                                                maxCallOutOutTimeAfterEnd = date.Add(ndtStartDate.TimeOfDay);
                                            }
                                        }
                                    }
                                }
                            }


                            if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                DateTime prevInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinInTime;
                                DateTime prevOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxOutTime;

                                DateTime prevCallOutInTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MinCallOutInTime;
                                DateTime prevCallOutOutTime = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()].MaxCallOutOutTime;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
                                        minInTime = prevInTime;
                                    }

                                    if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        maxOutTime = prevOutTime;
                                    }
                                    //}
                                    //else
                                    //{
                                    if (prevCallOutInTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                                    {
                                        if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }
                                    else
                                    {
                                        if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }

                                    if (prevCallOutOutTime.TimeOfDay < fdtWithAfterGraceTimeStartTime)
                                    {
                                        if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            maxCallOutOutTimeBeforeStart = prevCallOutOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                        {
                                            maxCallOutOutTimeAfterEnd = prevCallOutOutTime;
                                        }
                                    }


                                    //}
                                }
                                else
                                {
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
                                        minInTime = prevInTime;
                                    }

                                    if (maxOutTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        maxOutTime = prevOutTime;
                                    }
                                    //}
                                    //else
                                    //{
                                    if (prevCallOutInTime.TimeOfDay < ndtWithBeforeGraceTimeStartTime)
                                    {
                                        if (minCallOutInTimeBeforeStart.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }
                                    else
                                    {
                                        if (minCallOutInTimeAfterEnd.TimeOfDay > prevCallOutInTime.TimeOfDay)
                                        {
                                            minCallOutInTimeAfterEnd = prevCallOutInTime;
                                        }
                                    }

                                    if (prevCallOutOutTime.TimeOfDay < ndtWithAfterGraceTimeStartTime)
                                    {
                                        if (maxCallOutOutTimeBeforeStart.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            maxCallOutOutTimeBeforeStart = prevCallOutOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (maxCallOutOutTimeAfterEnd < prevCallOutOutTime)
                                        {
                                            maxCallOutOutTimeAfterEnd = prevCallOutOutTime;
                                        }
                                    }


                                    //}
                                }

                            }

                            int netNormalHours = 0;
                            int netNormalMinutes = 0;
                            int otHours = 0;
                            int otMinutes = 0;
                            int callOutHours = 0;
                            int callOutMinutes = 0;
                            string callOutFromHours = string.Empty;
                            string callOutToHours = string.Empty;
                            int lunchHours = 0;

                            inDateTimes.OrderBy((a) => a.TimeOfDay);
                            outDateTimes.OrderBy((a) => a.TimeOfDay);

                            foreach (DateTime inDateTime in inDateTimes)
                            {
                                //MessageBox.Show(this, "In Time: " + inDateTime.ToString());
                                DateTime outDateTime = DateTime.MaxValue;

                                //finding nearest out time wrt in time.
                                foreach (DateTime oDateTime in outDateTimes)
                                {
                                    if (oDateTime.TimeOfDay < inDateTime.TimeOfDay)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (oDateTime.TimeOfDay < outDateTime.TimeOfDay)
                                        {
                                            outDateTime = oDateTime;
                                        }
                                    }
                                }

                                //MessageBox.Show(this, "Out Time: " + outDateTime.ToString());

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    lunchHours = (fdtLunchEndTime - fdtLunchStartTime).Hours;
                                    //MessageBox.Show("Lunch Hours: " + lunchHours);
                                    //MessageBox.Show("In Hours: " + inTime.ToString());
                                    //MessageBox.Show("Out Hours: " + outTime.ToString());

                                    if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < fdtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours += (fdtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes += (fdtLunchStartTime - inDateTime.TimeOfDay).Minutes;
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours += (fdtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < fdtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > fdtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                                {
                                                    netNormalHours += (outDateTime.TimeOfDay - fdtLunchEndTime).Hours;
                                                    netNormalMinutes += (outDateTime.TimeOfDay - fdtLunchEndTime).Minutes;
                                                }
                                                else
                                                {
                                                    netNormalHours += (fdtEndTime - fdtLunchEndTime).Hours;
                                                    netNormalMinutes += (fdtEndTime - fdtLunchEndTime).Minutes;
                                                    otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                    otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= fdtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;
                                            }
                                            else
                                            {
                                                netNormalHours += (fdtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (fdtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - fdtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - fdtEndTime).Minutes;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    lunchHours = (ndtLunchEndTime - ndtLunchStartTime).Hours;

                                    //MessageBox.Show(this, "Lunch Hrs: " + lunchHours);

                                    if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchStartTime)
                                    {
                                        if (outDateTime.TimeOfDay < ndtWithAfterGraceTimeLunchEndTime)
                                        {
                                            netNormalHours += (ndtLunchStartTime - inDateTime.TimeOfDay).Hours;
                                            netNormalMinutes += (ndtLunchStartTime - inDateTime.TimeOfDay).Minutes;

                                            //MessageBox.Show(this, "ibl obl Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;

                                                //MessageBox.Show(this, "ibl oal obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }
                                            else
                                            {
                                                netNormalHours += (ndtEndTime - inDateTime.TimeOfDay).Hours - lunchHours;
                                                netNormalMinutes += (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                //MessageBox.Show(this, "ibl oal oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < ndtWithBeforeGraceTimeLunchEndTime)
                                        {
                                            if (outDateTime.TimeOfDay > ndtWithBeforeGraceTimeLunchEndTime)
                                            {
                                                if (outDateTime.TimeOfDay <= ndtWithBeforeGraceTimeEndTime)
                                                {
                                                    netNormalHours += (outDateTime.TimeOfDay - ndtLunchEndTime).Hours;
                                                    netNormalMinutes += (outDateTime.TimeOfDay - ndtLunchEndTime).Minutes;

                                                    //MessageBox.Show(this, "ible oale obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                                }
                                                else
                                                {
                                                    netNormalHours += (ndtEndTime - ndtLunchEndTime).Hours;
                                                    netNormalMinutes += (ndtEndTime - ndtLunchEndTime).Minutes;
                                                    otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                    otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                    //MessageBox.Show(this, "ible oale oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                                }
                                            }

                                        }
                                        else
                                        {
                                            if (outDateTime.TimeOfDay <= ndtWithAfterGraceTimeEndTime)
                                            {
                                                netNormalHours += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (outDateTime.TimeOfDay - inDateTime.TimeOfDay).Minutes;

                                                //MessageBox.Show(this, "iale obe Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }
                                            else
                                            {
                                                netNormalHours += (ndtEndTime - inDateTime.TimeOfDay).Hours;
                                                netNormalMinutes += (ndtEndTime - inDateTime.TimeOfDay).Minutes;
                                                otHours += (outDateTime.TimeOfDay - ndtEndTime).Hours;
                                                otMinutes += (outDateTime.TimeOfDay - ndtEndTime).Minutes;

                                                //MessageBox.Show(this, "iale oae Net hrs: " + netNormalHours + " Net Mins: " + netNormalMinutes);
                                            }

                                        }
                                    }
                                }
                            }

                            callOutInDateTimes.OrderBy((a) => a);
                            callOutOutDateTimes.OrderBy((a) => a);

                            foreach (DateTime callOutInDateTime in callOutInDateTimes)
                            {
                                DateTime callOutOutDateTime = DateTime.MaxValue;

                                //finding nearest out time wrt in time.
                                foreach (DateTime oDateTime in callOutOutDateTimes)
                                {
                                    if (oDateTime < callOutInDateTime)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (oDateTime < callOutOutDateTime)
                                        {
                                            callOutOutDateTime = oDateTime;
                                        }
                                    }
                                }

                                if (callOutInDateTime != DateTime.MaxValue && callOutOutDateTime != DateTime.MaxValue)
                                {
                                    callOutHours += (callOutOutDateTime - callOutInDateTime).Hours;
                                    callOutMinutes += (callOutOutDateTime - callOutInDateTime).Minutes;
                                }
                            }


                            if (minCallOutInTimeBeforeStart != DateTime.MaxValue && maxCallOutOutTimeBeforeStart != DateTime.MaxValue)
                            {
                                callOutFromHours = minCallOutInTimeBeforeStart.ToString("HH:mm");
                                callOutToHours = maxCallOutOutTimeBeforeStart.ToString("HH:mm");
                            }

                            if (minCallOutInTimeAfterEnd != DateTime.MaxValue && maxCallOutOutTimeAfterEnd != DateTime.MaxValue)
                            {
                                if (minCallOutInTimeBeforeStart == DateTime.MaxValue)
                                {
                                    callOutFromHours = minCallOutInTimeAfterEnd.ToString("HH:mm");
                                }

                                callOutToHours = maxCallOutOutTimeAfterEnd.ToString("HH:mm");
                            }

                            if (cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                CardHolderReportInfo reportInfo = cnicDateWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                if (reportInfo != null)
                                {
                                    reportInfo.NetNormalHours += netNormalHours;
                                    reportInfo.NetNormalMinutes += netNormalMinutes;
                                    reportInfo.OverTimeHours += otHours;
                                    reportInfo.OverTimeMinutes += otMinutes;
                                    reportInfo.TotalCallOutHours += callOutHours;
                                    reportInfo.TotalCallOutMinutes += callOutMinutes;

                                    if (minInTime.TimeOfDay < reportInfo.MinInTime.TimeOfDay)
                                    {
                                        reportInfo.MinInTime = minInTime;
                                    }

                                    if (maxOutTime.TimeOfDay > reportInfo.MaxOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxOutTime = maxOutTime;
                                    }

                                    if (minCallOutInTimeBeforeStart.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                    {
                                        reportInfo.MinCallOutInTime = minCallOutInTimeBeforeStart;
                                        reportInfo.CallOutFrom = callOutFromHours;
                                    }

                                    if (minCallOutInTimeAfterEnd.TimeOfDay < reportInfo.MinCallOutInTime.TimeOfDay)
                                    {
                                        reportInfo.MinCallOutInTime = minCallOutInTimeAfterEnd;
                                        reportInfo.CallOutFrom = callOutFromHours;
                                    }

                                    if (maxCallOutOutTimeAfterEnd.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxCallOutOutTime = maxCallOutOutTimeAfterEnd;
                                        reportInfo.CallOutTo = callOutToHours;
                                    }

                                    if (maxCallOutOutTimeAfterEnd == DateTime.MaxValue || maxCallOutOutTimeBeforeStart.TimeOfDay > reportInfo.MaxCallOutOutTime.TimeOfDay)
                                    {
                                        reportInfo.MaxCallOutOutTime = maxCallOutOutTimeBeforeStart;
                                        reportInfo.CallOutTo = callOutToHours;
                                    }
                                }
                            }
                            else
                            {
                                lstCnics.Add(cnicNumber);

                                cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                {
                                    OccurrenceTime = date,
                                    FirstName = chl.FirstName,
                                    PNumber = pNumber.ToString(),
                                    CNICNumber = cnicNumber,
                                    Department = department,
                                    Section = section,
                                    Cadre = cadre,
                                    NetNormalHours = netNormalHours,
                                    OverTimeHours = otHours,
                                    TotalCallOutHours = callOutHours,
                                    NetNormalMinutes = netNormalMinutes,
                                    OverTimeMinutes = otMinutes,
                                    TotalCallOutMinutes = callOutMinutes,
                                    CallOutFrom = callOutFromHours,
                                    CallOutTo = callOutToHours,
                                    MinInTime = minInTime,
                                    MaxOutTime = maxOutTime,
                                    MinCallOutInTime = minCallOutInTimeAfterEnd < minCallOutInTimeBeforeStart ? minCallOutInTimeAfterEnd : minCallOutInTimeBeforeStart,
                                    MaxCallOutOutTime = maxCallOutOutTimeAfterEnd == DateTime.MaxValue ? maxCallOutOutTimeBeforeStart : maxCallOutOutTimeAfterEnd
                                });





                            }

                            #endregion
                        }
                    }
                }

                if (progressBar1.InvokeRequired)
                {
                    progress = 12;
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress; }));
                }
                #endregion
              
                
                Dictionary<string, Dictionary<string, double>> CadreNic = new Dictionary<string, Dictionary<string, double>>();
                Dictionary<string, Dictionary<string, Dictionary<string, double>>> Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();



                if (cnicDateWiseReportInfo != null && cnicDateWiseReportInfo.Keys.Count > 0)
                {
                    int totalDays = (toDate.Date - fromDate.Date).Days;

                    for (int i = 0; i <= totalDays; i++)
                    {
                        DateTime date = fromDate.Date.AddDays(i);

                        foreach (string strCnic in lstCnics)
                        {
                            if (cnicDateWiseReportInfo.ContainsKey(strCnic + "^" + date.ToString()))
                            {
                                continue;
                            }
                            else
                            {
                                CardHolderReportInfo reportInfo = (from cnicDate in cnicDateWiseReportInfo
                                                                   where cnicDate.Key.Contains(strCnic)
                                                                   select cnicDate.Value).FirstOrDefault();

                                if (reportInfo != null)
                                {
                                    cnicDateWiseReportInfo.Add(strCnic + "^" + date.ToString(), new CardHolderReportInfo()
                                    {
                                        OccurrenceTime = date,
                                        FirstName = reportInfo.FirstName,
                                        PNumber = reportInfo.PNumber,
                                        CNICNumber = reportInfo.CNICNumber,
                                        Department = reportInfo.Department,
                                        Section = reportInfo.Section,
                                        Cadre = reportInfo.Cadre,
                                        MinInTime = DateTime.MaxValue,
                                        MaxOutTime = DateTime.MaxValue,
                                        MinCallOutInTime = DateTime.MaxValue,
                                        MaxCallOutOutTime = DateTime.MaxValue
                                    });


                                }
                            }
                        }
                    }


                    List<Cardholder> remainingCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                                             where chl != null &&
                                                                  !(from pds in chl.PersonalDataStrings
                                                                    where pds != null && pds.PersonalDataFieldID == 5051 && pds.Value != null && lstCnics.Contains(pds.Value)
                                                                    select pds).Any()
                                                             select chl).ToList();

                    foreach (Cardholder remainingChl in remainingCardHolders)
                    {
                        int pNumber = remainingChl.PersonalDataIntegers == null || remainingChl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(remainingChl.PersonalDataIntegers.ElementAt(0).Value);
                        string strPnumber = Convert.ToString(pNumber);
                        string cnicNumber = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051).Value);
                        string department = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043).Value);
                        string section = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951).Value);
                        string cadre = remainingChl.PersonalDataStrings == null ? string.Empty : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952) == null ? string.Empty : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952).Value);
                        string company = remainingChl.PersonalDataStrings == null ? "Unknown" : (remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059) == null ? "Unknown" : remainingChl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059).Value);

                        //Filter By Department
                        //if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() != filterByDepartment.ToLower())
                        //{
                        //    continue;
                        //}


                        //Filter By Section
                        if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                        {
                            continue;
                        }

                        //Filter By Cadre
                        if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && cadre.ToLower() != filterByCadre.ToLower())
                        {
                            continue;
                        }

                        //Filter By Company
                        if (!string.IsNullOrEmpty(filterByCompany) && company.ToLower() != filterByCompany.ToLower())
                        {
                            continue;
                        }

                        //Filter By CNIC
                        if (string.IsNullOrEmpty(cnicNumber) || !string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                        {
                            continue;
                        }

                        //Filter By Name
                        if (!string.IsNullOrEmpty(filerByName) && !remainingChl.FirstName.ToLower().Contains(filerByName.ToLower()))
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(filterByPnumber) && strPnumber != filterByPnumber)
                        {
                            continue;
                        }

                        for (int i = 0; i <= totalDays; i++)
                        {
                            DateTime date = fromDate.Date.AddDays(i);
                            if (!cnicDateWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                cnicDateWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                {
                                    OccurrenceTime = date,
                                    FirstName = remainingChl.FirstName,
                                    PNumber = strPnumber,
                                    CNICNumber = cnicNumber,
                                    Department = department,
                                    Section = section,
                                    Cadre = cadre,
                                    MinInTime = DateTime.MaxValue,
                                    MaxOutTime = DateTime.MaxValue,
                                    MinCallOutInTime = DateTime.MaxValue,
                                    MaxCallOutOutTime = DateTime.MaxValue
                                });




                            }

                        }
                    }
                    this.Summary_Report = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                    this.Depart_Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();
                    Dictionary<string, double> CnicHoursEffert = new Dictionary<string, double>();
                    //Dictionary<string, double> CnicHoursOthers = new Dictionary<string, double>();
                    int diff1 = 0;

                    int otherWorkers = 0;
                    int EffertWorkers = 0;

                    int NoEmpoyees = 0;
                    int attempt = 0;
                    foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicDateWiseReportInfo)
                    {

                        if (progress <= 80)
                        {
                            if (progressBar1.InvokeRequired)
                            {
                                progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = progress++; }));
                            }
                        }

                        if (reportInfo.Value == null)
                        {
                            continue;
                        }

                        string cnicNumber = reportInfo.Value.CNICNumber;
                        string department = reportInfo.Value.Department;
                        string section = reportInfo.Value.Section;
                        string cadre = reportInfo.Value.Cadre;



                        if (reportInfo.Value.Department == null)
                        {
                            continue;
                        }
                       
                        department = department.ToUpper();
                       

                        #region Making Hours


                        department = department.ToUpper();

                        if (Depart_Date_CadreNic.ContainsKey("PRODUCTION"))
                        {
                            if (Depart_Date_CadreNic["PRODUCTION"].ContainsKey("1/9/2020"))
                            {
                                if (Depart_Date_CadreNic["PRODUCTION"]["1/9/2020"].ContainsKey("EFFERT"))
                                {
                                    var eeem = Depart_Date_CadreNic["PRODUCTION"]["1/9/2020"]["EFFERT"]["NoOfEmployee"];
                                    if (eeem == 57)
                                    {
                                        attempt++;
                                    }
                                }
                            }
                        }

                       

                        if (reportInfo.Value.MinInTime != DateTime.MaxValue)
                        {
                            if (reportInfo.Value.MaxOutTime != DateTime.MaxValue)
                            {
                                diff1 = Convert.ToInt32((reportInfo.Value.MaxOutTime - reportInfo.Value.MinInTime).TotalHours);


                                if (CadreList.Contains(cadre))
                                {
                                    if (Depart_Date_CadreNic.ContainsKey(department))
                                    {
                                        if (Depart_Date_CadreNic[department].ContainsKey(reportInfo.Value.OccurrenceTime.Date.ToShortDateString()))
                                        {
                                            if (Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()].ContainsKey("EFFERT"))
                                            {
                                                if (Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"].ContainsKey("HoursCount"))
                                                {
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"]["HoursCount"] = Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"]["HoursCount"] + diff1;
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"]["NoOfEmployee"] = Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"]["NoOfEmployee"] + 1;


                                                    if (department.ToUpper() == "PRODUCTION")
                                                    {
                                                        NoEmpoyees++;
                                                    }
                                                }
                                                else
                                                {
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"].Add("HoursCount", diff1);
                                                     EffertWorkers = 0;
                                                     EffertWorkers++;
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["EFFERT"].Add("NoOfEmployee", otherWorkers);


                                                    if (department.ToUpper() == "PRODUCTION")
                                                    {
                                                        NoEmpoyees++;
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                CnicHoursEffert = new Dictionary<string, double>();
                                                CnicHoursEffert.Add("HoursCount", diff1);
                                                EffertWorkers = 0;
                                                EffertWorkers++;
                                                CnicHoursEffert.Add("NoOfEmployee", EffertWorkers);

                                                Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()].Add("EFFERT", CnicHoursEffert);



                                                if (department.ToUpper() == "PRODUCTION")
                                                {
                                                    NoEmpoyees++;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            CnicHoursEffert = new Dictionary<string, double>();
                                            CnicHoursEffert.Add("HoursCount", diff1);
                                            EffertWorkers = 0;
                                            EffertWorkers++;
                                            CnicHoursEffert.Add("NoOfEmployee", EffertWorkers);


                                            CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                            CadreNic.Add("EFFERT", CnicHoursEffert);

                                            Depart_Date_CadreNic[department].Add(reportInfo.Value.OccurrenceTime.Date.ToShortDateString(), CadreNic);


                                            if (department.ToUpper() == "PRODUCTION")
                                            {
                                                NoEmpoyees++;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        CnicHoursEffert = new Dictionary<string, double>();
                                        CnicHoursEffert.Add("HoursCount", diff1);
                                        EffertWorkers = 0;
                                        EffertWorkers++;
                                        CnicHoursEffert.Add("NoOfEmployee", EffertWorkers);


                                        CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                        CadreNic.Add("EFFERT", CnicHoursEffert);

                                        Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                                        Date_CadreNic.Add(reportInfo.Value.OccurrenceTime.Date.ToShortDateString(), CadreNic);

                                        Depart_Date_CadreNic.Add(department, Date_CadreNic);

                                        if (department.ToUpper() == "PRODUCTION")
                                        {
                                            NoEmpoyees++;
                                        }

                                    }



                                }
                                else
                                {
                                    if (Depart_Date_CadreNic.ContainsKey(department))
                                    {
                                        if (Depart_Date_CadreNic[department].ContainsKey(reportInfo.Value.OccurrenceTime.Date.ToShortDateString()))
                                        {
                                            if (Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()].ContainsKey("OTHERS"))
                                            {
                                                if (Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"].ContainsKey("HoursCount"))
                                                {
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"]["HoursCount"] = Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"]["HoursCount"] + diff1;
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"]["NoOfEmployee"] = Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"]["NoOfEmployee"] + 1;

                                                }
                                                else
                                                {
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"].Add("HoursCount", diff1);
                                                    otherWorkers++;
                                                    Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()]["OTHERS"].Add("NoOfEmployee", otherWorkers);
                                                }
                                            }
                                            else
                                            {
                                                CnicHoursEffert = new Dictionary<string, double>();
                                                CnicHoursEffert.Add("HoursCount", diff1);
                                                otherWorkers = 0;
                                                otherWorkers++;
                                                CnicHoursEffert.Add("NoOfEmployee", otherWorkers);

                                                Depart_Date_CadreNic[department][reportInfo.Value.OccurrenceTime.Date.ToShortDateString()].Add("OTHERS", CnicHoursEffert);

                                            }
                                        }
                                        else
                                        {
                                            CnicHoursEffert = new Dictionary<string, double>();
                                            CnicHoursEffert.Add("HoursCount", diff1);
                                            otherWorkers = 0;
                                            otherWorkers++;
                                            CnicHoursEffert.Add("NoOfEmployee", otherWorkers);


                                            CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                            CadreNic.Add("OTHERS", CnicHoursEffert);


                                            Depart_Date_CadreNic[department].Add(reportInfo.Value.OccurrenceTime.Date.ToShortDateString(), CadreNic);

                                        }
                                    }
                                    else
                                    {
                                        CnicHoursEffert = new Dictionary<string, double>();
                                        CnicHoursEffert.Add("HoursCount", diff1);
                                        otherWorkers = 0;
                                        otherWorkers++;
                                        CnicHoursEffert.Add("NoOfEmployee", otherWorkers);


                                        CadreNic = new Dictionary<string, Dictionary<string, double>>();
                                        CadreNic.Add("OTHERS", CnicHoursEffert);

                                        Date_CadreNic = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();
                                        Date_CadreNic.Add(reportInfo.Value.OccurrenceTime.Date.ToShortDateString(), CadreNic);

                                        Depart_Date_CadreNic.Add(department, Date_CadreNic);

                                    }




                                }
                            }
                        }

                        #endregion


                      
                    }
                }



                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = false; }));
                }
            }
            catch (Exception ex)
            {

                if (progressBar1.InvokeRequired)
                {
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Visible = false; }));
                    progressBar1.Invoke(new MethodInvoker(delegate { progressBar1.Value = 0; }));
                }
                string aa = ex.StackTrace.ToString();
                CreateLogFiles.ErrorLog(ex.Message + " " + ex.StackTrace);

                MessageBox.Show(ex.Message);
            }
        }
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);
            extension = extension.ToLower();
            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.Depart_Date_CadreNic, "Man-Hour Detail Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.Depart_Date_CadreNic, "Man-Hour Detail Report", "Man-Hour Detail Report");
            }
        }

        private void saveFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            string extension = Path.GetExtension(this.saveFileDialog2.FileName);
            extension = extension.ToLower();
            if (extension == ".pdf")
            {
                //this.SaveAsPdf1(this.Summary_Report, "Man-Hour Summary Report");
                this.SaveAsPdf1(this.Summary_Report, "Man-Hour Summary Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel1(this.Summary_Report, "Man-Hour Detail Report", "Man-Hour Summary Report");
            }

        }

        //For NEw Report



        private void SaveAsExcel(Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> data, string sheetName, string heading)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add(sheetName);

                        work.View.ShowGridLines = false;
                        work.Cells.Style.Font.Name = "Segoe UI Light";

                        work.Column(1).Width = 27;
                        work.Column(2).Width = 20;
                        work.Column(3).Width = 25;
                        work.Column(4).Width = 25;
                        work.Column(5).Width = 25;
                        work.Column(6).Width = 25;

                        //Heading
                        work.Cells["A1:B2"].Merge = true;
                        work.Cells["A1:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["A1:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:B2"].Style.Font.Size = 22;
                        work.Cells["A1:B2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        work.Cells["A1:B2"].Value = heading;

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 600);

                        int row = 4;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "From Date: ";
                        work.Cells[row, 2].Value = dtpFromDate.Value.Date.ToShortDateString();
                        work.Row(row).Height = 20;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "To Date: ";
                        work.Cells[row, 2].Value = dtpToDate.Value.Date.ToShortDateString();
                        work.Row(row).Height = 20;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;


                        //Sections and Data
                        string Department = "";
                        string Date = "";
                        int EffertNoofEmployee = 0;
                        int EffertHours = 0;
                        int OtherWorkers = 0;
                        int Othehours = 0;

                        int NetEffertNoofEmployee = 0;
                        int NetEffertHours = 0;
                        int NetOtherWorkers = 0;
                        int NetOthehours = 0;

                        row++;


                        work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                        work.Cells[row, 1, row, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;


                        work.Cells[row, 1].Value = "Date";
                        work.Cells[row, 2].Value = "Department";
                        work.Cells[row, 3].Value = "EFERT Employees Entry";
                        work.Cells[row, 4].Value = "Total Hours";
                        work.Cells[row, 5].Value = "Other Workers Entry";
                        work.Cells[row, 6].Value = "Total Hours";

                        work.Row(row).Height = 20;
                        var depardata = data[cbxDepartments.Text.ToUpper()];
                        Department = cbxDepartments.Text.ToUpper();
                        
                            foreach (var Date_CadreNics in depardata)
                            {
                                row++;
                                Date = Date_CadreNics.Key;

                                foreach (var CnicHoursEfferts in Date_CadreNics.Value)
                                {
                                    if (CnicHoursEfferts.Key == "EFFERT")
                                    {
                                        EffertHours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                        EffertNoofEmployee = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                        NetEffertHours += EffertHours;
                                        NetEffertNoofEmployee += EffertNoofEmployee;
                                    }
                                    else
                                    {
                                        Othehours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                        OtherWorkers = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                        NetOthehours += Othehours;
                                        NetOtherWorkers += OtherWorkers;
                                    }
                                }
                                work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                if (row % 2 == 0)
                                {
                                    work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                }

                                work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                work.Cells[row, 1].Value = Date;
                                work.Cells[row, 2].Value = Department;
                                work.Cells[row, 3].Value = EffertNoofEmployee;
                                work.Cells[row, 4].Value = EffertHours;
                                work.Cells[row, 5].Value = OtherWorkers;
                                work.Cells[row, 6].Value = Othehours;
                                work.Row(row).Height = 20;
                            }



                        row++;
                        work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Total";
                        work.Row(row).Style.Font.Bold = true;
                        work.Cells[row, 2].Value = "";
                        work.Cells[row, 3].Value = NetEffertNoofEmployee;
                        work.Cells[row, 4].Value = NetEffertHours;
                        work.Cells[row, 5].Value = NetOtherWorkers;
                        work.Cells[row, 6].Value = NetOthehours;
                        work.Row(row).Height = 20;

                        // FOr Manual Data 
                        row++;
                        work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Extra";
                       
                        work.Cells[row, 2].Value = "";
                        work.Cells[row, 3].Value = Manual_EffertNoofEmployee;
                        work.Cells[row, 4].Value = Manual_EffertHours;
                        work.Cells[row, 5].Value = Manual_OtherWorkers;
                        work.Cells[row, 6].Value = Manual_Othehours;
                        work.Row(row).Height = 20;


                        // Net Total

                        row++;
                        work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Net Total";
                        work.Row(row).Style.Font.Bold = true;
                        work.Cells[row, 2].Value = "";
                        work.Cells[row, 3].Value = NetEffertNoofEmployee + Manual_EffertNoofEmployee;
                        work.Cells[row, 4].Value = NetEffertHours + Manual_EffertHours;
                        work.Cells[row, 5].Value = NetOtherWorkers + Manual_OtherWorkers;
                        work.Cells[row, 6].Value = NetOthehours + Manual_Othehours;
                        work.Row(row).Height = 20;

                        // Manual_EffertNoofEmployee 
                        // Manual_EffertHours 
                        // Manual_OtherWorkers 
                        // Manual_Othehours 

                        ex.SaveAs(new System.IO.FileInfo(this.saveFileDialog1.FileName));

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.InnerException != null && exp.InnerException.InnerException != null)
                {
                    if (exp.InnerException.InnerException.HResult == -2147024864)
                    {
                        MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                    }
                    if (exp.InnerException.InnerException.HResult == -2147024891)
                    {
                        MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                    }
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }

        }

        private void SaveAsPdf(Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>> data, string heading)
        {
            Cursor currentCursor = Cursor.Current;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {
                    using (PdfWriter pdfWriter = new PdfWriter(this.saveFileDialog1.FileName))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(pdfWriter))
                        {
                            using (Document doc = new Document(pdfDocument))
                            {
                                doc.SetFont(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"));
                                string headerLeftText = "Report From: " + this.dtpFromDate.Value.ToShortDateString() + " To: " + this.dtpToDate.Value.ToShortDateString();
                                string headerRightText = string.Empty;
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString();

                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1400F, 842F));
                                Table table = new Table((new List<float>() { 70F, 150F, 100F, 150F, 220F, 120F, 150F, 220F, 120F }).ToArray());

                                table.SetWidth(1300F);
                                table.SetFixedLayout();
                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                this.AddTableHeaderRow(table);

                                string Department = "";
                                string Date = "";


                                int NetEffertNoofEmployee = 0;
                                int NetEffertHours = 0;
                                int NetOtherWorkers = 0;
                                int NetOthehours = 0;


                                int i = 0;


                                var depardata = data[cbxDepartments.Text.ToUpper()];
                                Department = cbxDepartments.Text.ToUpper();
                                int EffertNoofEmployee = 0;
                                int EffertHours = 0;
                                int OtherWorkers = 0;
                                int Othehours = 0;
                                foreach (var Date_CadreNics in depardata)
                                {
                                    EffertNoofEmployee = 0;
                                    EffertHours = 0;
                                    OtherWorkers = 0;
                                    Othehours = 0;
                                    i++;
                                    Date = Date_CadreNics.Key;

                                    foreach (var CnicHoursEfferts in Date_CadreNics.Value)
                                    {
                                        if (CnicHoursEfferts.Key == "EFFERT")
                                        {
                                            EffertHours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                            EffertNoofEmployee = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                            NetEffertHours += EffertHours;
                                            NetEffertNoofEmployee += EffertNoofEmployee;
                                        }
                                        else
                                        {
                                            Othehours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                            OtherWorkers = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                            NetOthehours += Othehours;
                                            NetOtherWorkers += OtherWorkers;
                                        }
                                    }

                                    this.AddTableDataRow(table, Date, Department, EffertNoofEmployee.ToString(), EffertHours.ToString(), OtherWorkers.ToString(), Othehours.ToString(), i % 2 == 0);

                                }







                                this.AddTableDataRow(table, "Total :", "", NetEffertNoofEmployee.ToString(), NetEffertHours.ToString(), NetOtherWorkers.ToString(), NetOthehours.ToString(), i % 2 == 0);

                                this.AddTableDataRow(table, "Extra :", "", Manual_EffertNoofEmployee.ToString(), Manual_EffertHours.ToString(), Manual_OtherWorkers.ToString(), Manual_Othehours.ToString(), i % 2 != 0);

                                this.AddTableDataRow(table, "Net Total :", "", (NetEffertNoofEmployee + Manual_EffertNoofEmployee).ToString(), (NetEffertHours + Manual_EffertHours).ToString(), (NetOtherWorkers + Manual_OtherWorkers).ToString(), (NetOthehours + Manual_Othehours).ToString(), i % 2 == 0);



                                doc.Add(table);

                                doc.Close();
                            }
                        }

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                CreateLogFiles.ErrorLog(exp.Message + " " + exp.StackTrace);
                Cursor.Current = currentCursor;
                if (exp.HResult == -2147024864)
                {
                    MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                }
                else
                if (exp.HResult == -2147024891)
                {
                    MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }
        }

        private void SaveAsPdf1(Dictionary<string, Dictionary<string, Dictionary<string, double>>> data, string heading)
        {
            Cursor currentCursor = Cursor.Current;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {
                    using (PdfWriter pdfWriter = new PdfWriter(this.saveFileDialog2.FileName))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(pdfWriter))
                        {
                            using (Document doc = new Document(pdfDocument))
                            {
                                doc.SetFont(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"));
                                string headerLeftText = "Report From: " + this.dtpFromDate.Value.ToShortDateString() + " To: " + this.dtpToDate.Value.ToShortDateString();
                                string headerRightText = string.Empty;
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString();

                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1400F, 842F));
                                Table table = new Table((new List<float>() { 70F, 150F, 100F, 150F, 220F, 120F, 150F, 220F, 120F }).ToArray());

                                table.SetWidth(1300F);
                                table.SetFixedLayout();
                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                this.AddTableHeaderRow(table);

                                string Department = "";
                                string Date = "";
                                

                                int NetEffertNoofEmployee = 0;
                                int NetEffertHours = 0;
                                int NetOtherWorkers = 0;
                                int NetOthehours = 0;


                                int i = 0;
                                foreach (var Depart in data)
                                {
                                    int EffertNoofEmployee = 0;
                                    int EffertHours = 0;
                                    int OtherWorkers = 0;
                                    int Othehours = 0;
                                    Date = (i++).ToString();
                                    Department = Depart.Key.ToUpper();

                                    
                                    foreach (var CnicHoursEfferts in Depart.Value)
                                    {
                                        if (CnicHoursEfferts.Key == "EFFERT")
                                        {
                                            EffertHours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                            EffertNoofEmployee = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                            NetEffertHours += EffertHours;
                                            NetEffertNoofEmployee += EffertNoofEmployee;
                                        }
                                        else
                                        {
                                            Othehours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                            OtherWorkers = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                            NetOthehours += Othehours;
                                            NetOtherWorkers += OtherWorkers;
                                        }


                                    }
                                    this.AddTableDataRow(table, Date, Department, EffertNoofEmployee.ToString(), EffertHours.ToString(), OtherWorkers.ToString(), Othehours.ToString(), i % 2 == 0);

                                }


                                this.AddTableDataRow(table, "Total :", "", NetEffertNoofEmployee.ToString(), NetEffertHours.ToString(), NetOtherWorkers.ToString(), NetOthehours.ToString(), i % 2 == 0);

                                this.AddTableDataRow(table, "Extra :", "", Manual_EffertNoofEmployee.ToString(), Manual_EffertHours.ToString(), Manual_OtherWorkers.ToString(), Manual_Othehours.ToString(), i % 2 == 0);

                                this.AddTableDataRow(table, "Net Total :", "", (NetEffertNoofEmployee + Manual_EffertNoofEmployee).ToString(), (NetEffertHours + Manual_EffertHours).ToString(), (NetOtherWorkers + Manual_OtherWorkers).ToString(), (NetOthehours + Manual_Othehours).ToString(), i % 2 == 0);



                                doc.Add(table);

                                doc.Close();
                            }
                        }

                        System.Diagnostics.Process.Start(this.saveFileDialog2.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                CreateLogFiles.ErrorLog(exp.Message + " " + exp.StackTrace);
                Cursor.Current = currentCursor;
                if (exp.HResult == -2147024864)
                {
                    MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                }
                else
                if (exp.HResult == -2147024891)
                {
                    MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }
        }


        private void SaveAsExcel1(Dictionary<string, Dictionary<string, Dictionary<string, double>>> data, string sheetName, string heading)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add(sheetName);

                        work.View.ShowGridLines = false;
                        work.Cells.Style.Font.Name = "Segoe UI Light";

                        work.Column(1).Width = 35;
                        work.Column(2).Width = 30;
                        work.Column(3).Width = 25;
                        work.Column(4).Width = 25;
                        work.Column(5).Width = 25;
                        //work.Column(6).Width = 25;

                        //Heading
                        

                        work.Cells["A1:B2"].Merge = true;
                     
                        work.Cells["A1:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["A1:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:B2"].Style.Font.Size = 22;
                        work.Cells["A1:B2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        work.Cells["A1:B2"].Value = heading;
                        
                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 600);

                        int row = 4;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "From Date: ";
                        work.Cells[row, 2].Value = dtpFromDate.Value.Date.ToShortDateString();
                        work.Row(row).Height = 20;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "To Date: ";
                        work.Cells[row, 2].Value = dtpToDate.Value.Date.ToShortDateString();
                        work.Row(row).Height = 20;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;


                        //Sections and Data
                        string Department = "";
                      
                        int EffertNoofEmployee = 0;
                        int EffertHours = 0;
                        int OtherWorkers = 0;
                        int Othehours = 0;

                        int NetEffertNoofEmployee = 0;
                        int NetEffertHours = 0;
                        int NetOtherWorkers = 0;
                        int NetOthehours = 0;

                        row++;


                        work.Cells[row, 1, row, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 5].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        work.Cells[row, 1, row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                        work.Cells[row, 1, row, 5].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;


                        //work.Cells[row, 1].Value = "Date";
                        work.Cells[row, 1].Value = "Department";
                        work.Cells[row, 2].Value = "EFERT Employees Entry";
                        work.Cells[row, 3].Value = "Total Hours";
                        work.Cells[row, 4].Value = "Other Workers Entry";
                        work.Cells[row, 5].Value = "Total Hours";

                        work.Row(row).Height = 20;

                        foreach (var Depart in data)
                        {
                            row++;
                            Department = Depart.Key.ToUpper();

                            foreach (var CnicHoursEfferts in Depart.Value)
                            {
                                if (CnicHoursEfferts.Key == "EFFERT")
                                {
                                    EffertHours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                    EffertNoofEmployee = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                    NetEffertHours += EffertHours;
                                    NetEffertNoofEmployee += EffertNoofEmployee;
                                }
                                else
                                {
                                    Othehours = Convert.ToInt32(CnicHoursEfferts.Value["HoursCount"].ToString());
                                    OtherWorkers = Convert.ToInt32(CnicHoursEfferts.Value["NoOfEmployee"].ToString());

                                    NetOthehours += Othehours;
                                    NetOtherWorkers += OtherWorkers;
                                }
                            }
                            work.Cells[row, 1, row, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                            work.Cells[row, 1, row, 5].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 5].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 5].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 5].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                            if (row % 2 == 0)
                            {
                                work.Cells[row, 1, row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                work.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                            }

                            work.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            //work.Cells[row, 1].Value = Date;
                            work.Cells[row, 1].Value = Department;
                            work.Cells[row, 2].Value = EffertNoofEmployee;
                            work.Cells[row, 3].Value = EffertHours;
                            work.Cells[row, 4].Value = OtherWorkers;
                            work.Cells[row, 5].Value = Othehours;
                            work.Row(row).Height = 20;




                        }

                        row++;
                        work.Cells[row, 1, row, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 5].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Total";
                        work.Row(row).Style.Font.Bold = true;
                        
                        work.Cells[row, 2].Value = NetEffertNoofEmployee;
                        work.Cells[row, 3].Value = NetEffertHours;
                        work.Cells[row, 4].Value = NetOtherWorkers;
                        work.Cells[row, 5].Value = NetOthehours;
                        work.Row(row).Height = 20;

                        // FOr Manual Data 
                        row++;
                        work.Cells[row, 1, row, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 5].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Extra";

                        work.Cells[row, 2].Value = Manual_EffertNoofEmployee;
                        work.Cells[row, 3].Value = Manual_EffertHours;
                        work.Cells[row, 4].Value = Manual_OtherWorkers;
                        work.Cells[row, 5].Value = Manual_Othehours;
                        work.Row(row).Height = 20;


                        // Net Total

                        row++;
                        work.Cells[row, 1, row, 5].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1, row, 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        work.Cells[row, 1, row, 5].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 5].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                        if (row % 2 == 0)
                        {
                            work.Cells[row, 1, row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        work.Cells[row, 1, row, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        work.Cells[row, 1].Value = "Net Total";
                        work.Row(row).Style.Font.Bold = true;
                       
                        work.Cells[row, 2].Value = NetEffertNoofEmployee + Manual_EffertNoofEmployee;
                        work.Cells[row, 3].Value = NetEffertHours + Manual_EffertHours;
                        work.Cells[row, 4].Value = NetOtherWorkers + Manual_OtherWorkers;
                        work.Cells[row, 5].Value = NetOthehours + Manual_Othehours;
                        work.Row(row).Height = 20;

                        // Manual_EffertNoofEmployee 
                        // Manual_EffertHours 
                        // Manual_OtherWorkers 
                        // Manual_Othehours 

                        ex.SaveAs(new System.IO.FileInfo(this.saveFileDialog2.FileName));

                        System.Diagnostics.Process.Start(this.saveFileDialog2.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.InnerException != null && exp.InnerException.InnerException != null)
                {
                    if (exp.InnerException.InnerException.HResult == -2147024864)
                    {
                        MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                    }
                    if (exp.InnerException.InnerException.HResult == -2147024891)
                    {
                        MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                    }
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }

        }


        private void AddMainHeading(Table table, string heading)
        {
            Cell headingCell = new Cell(2, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
                // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(80F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
        }

        private void AddNewEmptyRow(Table table, bool removeBottomBorder = true)
        {
            table.StartNewRow();

            if (removeBottomBorder)
            {
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }
            else
            {
                table.AddCell(new Cell().
                   SetHeight(6F).
                   SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                   SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                   SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));


            }

        }

        private void AddTableHeaderRow(Table table)
        {
            table.StartNewRow();

            if (reportType.Text.ToString() == "Man-Hours Detail Report")
            {
                table.AddCell(new Cell().
                   Add(new Paragraph("Date").
                   SetFontSize(11F)).
               SetBackgroundColor(new DeviceRgb(253, 233, 217)).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            }
            else
            {
                table.AddCell(new Cell().
                  Add(new Paragraph("S.NO").
                  SetFontSize(11F)).
              SetBackgroundColor(new DeviceRgb(253, 233, 217)).
              SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
              SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
              SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            }

            table.AddCell(new Cell().
                    Add(new Paragraph("Department").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                    Add(new Paragraph("EFERT Employees Entry").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Total Hours").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Other Workers Entry").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Total Hours").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
           }

        private void AddTableDataRow(Table table, string Date, string Department, string EFERT_Employees_Entry, string Total_Hours, string Other_Workers_Entry, string Total_Hours_o, bool altRow)
        {
            


            table.StartNewRow();
            table.AddCell(new Cell().
                   Add(new Paragraph(string.IsNullOrEmpty(Date) ? string.Empty : Date).
                   SetFontSize(11F)).
               SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(Department) ? string.Empty : Department).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE))
                ;
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(EFERT_Employees_Entry) ? string.Empty : EFERT_Employees_Entry).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(Total_Hours) ? string.Empty : Total_Hours).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(Other_Workers_Entry) ? string.Empty : Other_Workers_Entry).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(Total_Hours_o).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            
            
        }

    
    }
}
