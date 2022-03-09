using AttendanceReport.CCFTCentral;
using AttendanceReport.CCFTEvent;
using AttendanceReport.EFERTDb;
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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class Form1 : Form
    {
        
        List<CardHolderReportInfo> lstCardHolderReportInfo = null;
        public Form1()
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFieldsMultiCombobox(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre, this.cbxCrew);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {                

                Cursor.Current = Cursors.WaitCursor;
                
                lstCardHolderReportInfo = new List<CardHolderReportInfo>();

                DateTime fromDate = this.dtpFromDate.Value.Date;
                DateTime fromDateUtc = fromDate.ToUniversalTime();
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                DateTime toDateUtc = toDate.ToUniversalTime();

                Dictionary<string, CardHolderReportInfo> cnicWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();

                List<string> checkinBeforeStartTimeLst = new List<string>();

              
                string filterByDepartment = this.cbxDepartments.Text;
                string filterBySection = this.cbxSections.Text;
                string filerByName = this.tbxName.Text;
                string filterByPnumber = this.tbxPNumber.Text;
                string filterByCardNumber = this.tbxCarNumber.Text;
                string filterByCrew = this.cbxCrew.Text;
                string filterByCadre = this.cbxCadre.Text;
                string filterByCompany = this.cbxCompany.Text;
                string filterByCNIC = this.tbxCnic.Text;

                TimeSpan thStartTime = this.dtpLateTimeStart.Value.TimeOfDay;
                TimeSpan thEndTime = this.dtpLateTimeEnd.Value.TimeOfDay;

                List<CCFTEvent.Event> lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                                   where
                                                       events != null && (events.EventType == 20001) &&
                                                       events.OccurrenceTime >= fromDateUtc &&
                                                       events.OccurrenceTime < toDateUtc
                                                   select events).ToList();

                //MessageBox.Show(this, "Events Found:" + lstEvents.Count);

                #region Dummy Events

                //List<CCFTEvent.Event> lstEvents = new List<CCFTEvent.Event>()
                //{
                //         new CCFTEvent.Event() {
                //            EventType = 20001,
                //            OccurrenceTime = new DateTime(2019,10,11,09,50,44,DateTimeKind.Utc),
                //            RelatedItems = new List<RelatedItem>() {
                //                new RelatedItem() {
                //                    RelationCode = 0,
                //                    FTItemID = 1046
                //                }
                //            }
                //        },
                //             new CCFTEvent.Event() {
                //            EventType = 20001,
                //            OccurrenceTime = new DateTime(2019,10,10,09,40,44,DateTimeKind.Utc),
                //            RelatedItems = new List<RelatedItem>() {
                //                new RelatedItem() {
                //                    RelationCode = 0,
                //                    FTItemID = 1046
                //                }
                //            }
                //        },
                //                 new CCFTEvent.Event() {
                //            EventType = 20001,
                //            OccurrenceTime = new DateTime(2019,10,07,09,45,44,DateTimeKind.Utc),
                //            RelatedItems = new List<RelatedItem>() {
                //                new RelatedItem() {
                //                    RelationCode = 0,
                //                    FTItemID = 1046
                //                }
                //            }
                //        },
                //                     new CCFTEvent.Event() {
                //            EventType = 20001,
                //            OccurrenceTime = new DateTime(2019,10,07,08,43,44,DateTimeKind.Utc),
                //            RelatedItems = new List<RelatedItem>() {
                //                new RelatedItem() {
                //                    RelationCode = 0,
                //                    FTItemID = 1046
                //                }
                //            }
                //        },
                //           new CCFTEvent.Event() {
                //            EventType = 20001,
                //            OccurrenceTime = new DateTime(2019,10,07,09,42,44,DateTimeKind.Utc),
                //            RelatedItems = new List<RelatedItem>() {
                //                new RelatedItem() {
                //                    RelationCode = 0,
                //                    FTItemID = 14716
                //                }
                //            }
                //        },
                //    new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,10,07,09,43,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 10864
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,10,08,09,03,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 8288
                //            }
                //        }
                //    },
                    // new CCFTEvent.Event() {
                    //    EventType = 20001,
                    //    OccurrenceTime = new DateTime(2018,09,16,10,15,09,DateTimeKind.Utc),
                    //    RelatedItems = new List<RelatedItem>() {
                    //        new RelatedItem() {
                    //            RelationCode = 0,
                    //            FTItemID = 1046
                    //        }
                    //    }
                    //},
                    //   new CCFTEvent.Event() {
                    //    EventType = 20003,
                    //    OccurrenceTime = new DateTime(2018,09,16,11,58,09,DateTimeKind.Utc),
                    //    RelatedItems = new List<RelatedItem>() {
                    //        new RelatedItem() {
                    //            RelationCode = 0,
                    //            FTItemID = 1046
                    //        }
                    //    }
                    //},
                    //    new CCFTEvent.Event() {
                    //    EventType = 20001,
                    //    OccurrenceTime = new DateTime(2018,09,16,1,15,09,DateTimeKind.Utc),
                    //    RelatedItems = new List<RelatedItem>() {
                    //        new RelatedItem() {
                    //            RelationCode = 0,
                    //            FTItemID = 1046
                    //        }
                    //    }
                    //},
                    //   new CCFTEvent.Event() {
                    //    EventType = 20003,
                    //    OccurrenceTime = new DateTime(2018,09,16,1,58,09,DateTimeKind.Utc),
                    //    RelatedItems = new List<RelatedItem>() {
                    //        new RelatedItem() {
                    //            RelationCode = 0,
                    //            FTItemID = 1046
                    //        }
                    //    }
                    //},
                     //};

                    #endregion



                    List<int> inIds = new List<int>();
                //date and ftitemid of person an thier event entries
                Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();

                Dictionary<int, Cardholder> inCardHolders = new Dictionary<int, Cardholder>();

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
                            if (relatedItem.RelationCode == 0)//human
                            {
                                if (events.EventType == 20001)//In Events
                                {
                                    inIds.Add(relatedItem.FTItemID);

                                    if (lstChlEvents.ContainsKey(events.OccurrenceTime.Date))
                                    {
                                        if (lstChlEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                        {
                                            lstChlEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);

                                        }
                                        else
                                        {

                                            lstChlEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                        }
                                    }
                                    else
                                    {
                                        dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                        dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                        lstChlEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                    }
                                }
                            }

                        }
                    }
                }


                inCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                                 where chl != null && inIds.Contains(chl.FTItemID)
                                 select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);

                //MessageBox.Show(this, "In CHls Found Keys: " + inCardHolders.Keys.Count + " Values: " + inCardHolders.Values.Count);

                List<string> strLstTempCards = (from chl in inCardHolders
                                                where chl.Value != null && (chl.Value.FirstName.ToLower().StartsWith("t-") || chl.Value.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-"))
                                                select chl.Value.LastName).ToList();

            //MessageBox.Show(this, "Temp Cards found: " + strLstTempCards.Count);

            List<CheckInAndOutInfo> filteredCheckIns = (from checkin in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                                        where checkin != null && !checkin.CheckedIn && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate
                                                        && strLstTempCards.Contains(checkin.CardNumber)
                                                        select checkin).ToList();

                if (!string.IsNullOrEmpty(filterByCNIC))
                {
                    filteredCheckIns = (from checkin in filteredCheckIns
                                        where checkin != null && ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.CNICNumber == filterByCNIC) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.CNICNumber == filterByCNIC) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.CNICNumber == filterByCNIC))
                                        select checkin).ToList();
                }
                else if (!string.IsNullOrEmpty(filerByName))
                {
                    filerByName = filerByName.ToLower();
                    filteredCheckIns = (from checkin in filteredCheckIns
                                        where checkin != null && ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.FirstName.ToLower().Contains(filerByName)) ||
                                                                    (checkin.Visitors != null &&
                                                                    checkin.Visitors.FirstName.ToLower().Contains(filerByName)))
                                        select checkin).ToList();
                }
                else if (!string.IsNullOrEmpty(filterByCardNumber))
                {
                    filteredCheckIns = (from checkin in filteredCheckIns
                                        where checkin != null && checkin.CardHolderInfos != null && checkin.CardHolderInfos.CardNumber == filterByCardNumber
                                        select checkin).ToList();
                }
                else if (!string.IsNullOrEmpty(filterByPnumber))
                {
                    filteredCheckIns = (from checkin in filteredCheckIns
                                        where checkin != null && checkin.CardHolderInfos != null && checkin.CardHolderInfos.PNumber == filterByPnumber
                                        select checkin).ToList();
                }
                else
                {

                    List<CheckInAndOutInfo> filteredCheckInsNew = new List<CheckInAndOutInfo>();

                    for (int i = 0; i < filteredCheckIns.Count; i++)
                    {
                        CheckInAndOutInfo checkInAndOutInfo = filteredCheckIns[i];

                        if (checkInAndOutInfo == null)
                        {
                            continue;
                        }

                        //filterBySection
                        if (!string.IsNullOrEmpty(filterBySection))
                        {

                            string section = string.Empty;
                            if (checkInAndOutInfo.CardHolderInfos != null && checkInAndOutInfo.CardHolderInfos.Section != null)
                            {
                                section = checkInAndOutInfo.CardHolderInfos.Section.SectionName;
                            }
                            else
                            {
                                if (checkInAndOutInfo.DailyCardHolders != null && !string.IsNullOrEmpty(checkInAndOutInfo.DailyCardHolders.Section))
                                {
                                    section = checkInAndOutInfo.DailyCardHolders.Section;
                                }
                            }

                            if (!string.IsNullOrEmpty(section))
                            {
                                bool isValidEntry = this.isValidEntry(filterBySection, section);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }
                        }

                        //filterByCadre
                        if (!string.IsNullOrEmpty(filterByCadre))
                        {

                            string cadre = string.Empty;
                            if (checkInAndOutInfo.CardHolderInfos != null && checkInAndOutInfo.CardHolderInfos.Cadre != null)
                            {
                                cadre = checkInAndOutInfo.CardHolderInfos.Cadre.CadreName;
                            }
                            else
                            {
                                if (checkInAndOutInfo.DailyCardHolders != null && !string.IsNullOrEmpty(checkInAndOutInfo.DailyCardHolders.Cadre))
                                {
                                    cadre = checkInAndOutInfo.DailyCardHolders.Cadre;
                                }
                            }

                            if (!string.IsNullOrEmpty(cadre))
                            {
                                bool isValidEntry = this.isValidEntry(filterByCadre, cadre);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }
                        }

                        //filterByCrew
                        if (!string.IsNullOrEmpty(filterByCrew))
                        {

                            string crew = string.Empty;
                            if (checkInAndOutInfo.CardHolderInfos != null && checkInAndOutInfo.CardHolderInfos.Crew != null)
                            {
                                crew = checkInAndOutInfo.CardHolderInfos.Crew.CrewName;
                            }
                           
                            if (!string.IsNullOrEmpty(crew))
                            {
                                bool isValidEntry = this.isValidEntry(filterByCrew, crew);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }
                        }

                        //filterByDepartment
                        if (!string.IsNullOrEmpty(filterByDepartment))
                        {

                            string dept = string.Empty;
                            if (checkInAndOutInfo.CardHolderInfos != null && checkInAndOutInfo.CardHolderInfos.Department != null)
                            {
                                dept = checkInAndOutInfo.CardHolderInfos.Department.DepartmentName;
                            }
                            else
                            {
                                if (checkInAndOutInfo.DailyCardHolders != null && !string.IsNullOrEmpty(checkInAndOutInfo.DailyCardHolders.Department))
                                {
                                    dept = checkInAndOutInfo.DailyCardHolders.Department;
                                }
                            }

                            if (!string.IsNullOrEmpty(dept))
                            {
                                bool isValidEntry = this.isValidEntry(filterByDepartment, dept);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }
                        }

                        //filterByCompany
                        if (!string.IsNullOrEmpty(filterByCompany))
                        {
                            string company = string.Empty;
                            if (checkInAndOutInfo.CardHolderInfos != null && checkInAndOutInfo.CardHolderInfos.Company != null)
                            {
                                company = checkInAndOutInfo.CardHolderInfos.Company.CompanyName;
                            }
                            else if (checkInAndOutInfo.DailyCardHolders != null && !string.IsNullOrEmpty(checkInAndOutInfo.DailyCardHolders.CompanyName))
                            {
                                company = checkInAndOutInfo.DailyCardHolders.CompanyName;
                            }
                            else
                            {
                                if (checkInAndOutInfo.Visitors != null && !string.IsNullOrEmpty(checkInAndOutInfo.Visitors.CompanyName))
                                {
                                    company = checkInAndOutInfo.Visitors.CompanyName;
                                }
                            }

                            if (!string.IsNullOrEmpty(company))
                            {
                                bool isValidEntry = this.isValidEntry(filterByCompany, company);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }
                        }


                        filteredCheckInsNew.Add(checkInAndOutInfo);
                    }

                    filteredCheckIns = filteredCheckInsNew;
                }
                //MessageBox.Show(this, "Filtered Checkins: " + filteredCheckIns.Count);

                foreach (KeyValuePair<DateTime, Dictionary<int, List<CCFTEvent.Event>>> inEvent in lstChlEvents)
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

                            Dictionary<string, DateTime> dictCallOutInTime = new Dictionary<string, DateTime>();
                            Dictionary<string, DateTime> dictCallOutOutTime = new Dictionary<string, DateTime>();

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


                                DateTime inDateTime = DateTime.MaxValue;
                                DateTime outDateTime = DateTime.MaxValue;



                                if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                {
                                    CardHolderReportInfo reportInfo = cnicWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                    if (dateWiseCheckIn.DateTimeIn.TimeOfDay < reportInfo.OccurrenceTime.TimeOfDay)
                                    {
                                        if (TimeSpan.Compare(dateWiseCheckIn.DateTimeIn.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(dateWiseCheckIn.DateTimeIn.TimeOfDay, thEndTime) <= 0)
                                        {
                                            reportInfo.CardNumber = chl.LastName;
                                            reportInfo.OccurrenceTime = dateWiseCheckIn.DateTimeIn;
                                        }
                                        else
                                        {
                                            cnicWiseReportInfo.Remove(cnicNumber + "^" + date.ToString());
                                        }
                                    }


                                }
                                else
                                {
                                    if (TimeSpan.Compare(dateWiseCheckIn.DateTimeIn.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(dateWiseCheckIn.DateTimeIn.TimeOfDay, thEndTime) <= 0)
                                    {
                                        cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                        {
                                            CardNumber = chl.LastName,
                                            OccurrenceTime = dateWiseCheckIn.DateTimeIn,
                                            FirstName = chl.FirstName,
                                            PNumber = pNumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre
                                        });
                                    }
                                    else
                                    {
                                        //user checkin before start time range list
                                        if (TimeSpan.Compare(dateWiseCheckIn.DateTimeIn.TimeOfDay, thStartTime) < 0)
                                        {
                                            checkinBeforeStartTimeLst.Add(cnicNumber + "^" + date.ToString());
                                        }
                                    }
                                }
                                
                            }

                            #endregion
                        }
                        else
                        {
                            #region Events


                            List<CCFTEvent.Event> events = chlWiseEvents.Value;

                            events = events.OrderBy(ev => ev.OccurrenceTime).ToList();


                            int pNumber = chl.PersonalDataIntegers == null || chl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(chl.PersonalDataIntegers.ElementAt(0).Value);
                            string strPnumber = Convert.ToString(pNumber);
                            string cnicNumber = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5051).Value);
                            string department = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5043).Value);
                            string section = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12951).Value);
                            string cadre = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12952).Value);
                            string company = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 5059).Value);
                            string crew = chl.PersonalDataStrings == null ? "Unknown" : (chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12869) == null ? "Unknown" : chl.PersonalDataStrings.ToList().Find(pds => pds.PersonalDataFieldID == 12869).Value);

                            strPnumber = string.IsNullOrEmpty(strPnumber) ? "Unknown" : strPnumber;
                            cnicNumber = string.IsNullOrEmpty(cnicNumber) ? "Unknown" : cnicNumber;
                            department = string.IsNullOrEmpty(department) ? "Unknown" : department;
                            section = string.IsNullOrEmpty(section) ? "Unknown" : section;
                            cadre = string.IsNullOrEmpty(cadre) ? "Unknown" : cadre;
                            company = string.IsNullOrEmpty(company) ? "Unknown" : company;
                            crew = string.IsNullOrEmpty(crew) ? "Unknown" : crew;

                            //Filter By CNIC 31303-4961345-5
                            if (!string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                            {
                                continue;
                            }

                            //Filter By Name
                            if (!string.IsNullOrEmpty(filerByName) && !chl.FirstName.ToLower().Contains(filerByName.ToLower()))
                            {
                                continue;
                            }

                            //Filter By Card Number
                            if (!string.IsNullOrEmpty(filterByCardNumber))
                            {
                                int cardNumber;

                                bool parsed = Int32.TryParse(chl.LastName, out cardNumber);

                                if (parsed)
                                {
                                    if (cardNumber != Convert.ToInt32(filterByCardNumber))
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    continue;
                                }

                            }

                            if (!string.IsNullOrEmpty(filterByPnumber) && strPnumber != filterByPnumber)
                            {
                                continue;
                            }




                            //Filter By Section
                            if (!string.IsNullOrEmpty(filterBySection))
                            {
                                bool isValidEntry = this.isValidEntry(filterBySection, section);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }


                            //Filter By Cadre
                            if (!string.IsNullOrEmpty(filterByCadre))
                            {

                                bool isValidEntry = this.isValidEntry(filterByCadre, cadre);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }



                            //Filter By Crew
                            if (!string.IsNullOrEmpty(filterByCrew))
                            {
                                bool isValidEntry = this.isValidEntry(filterByCrew, crew);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }


                            //Filter By Department
                            if (!string.IsNullOrEmpty(filterByDepartment))
                            {
                                bool isValidEntry = this.isValidEntry(filterByDepartment, department);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }

                            //Filter By Company
                            if (!string.IsNullOrEmpty(filterByCompany))
                            {
                                bool isValidEntry = this.isValidEntry(filterByCompany, company);
                                if (!isValidEntry)
                                {
                                    continue;
                                }
                            }



                            DateTime minInTime = DateTime.MaxValue;
                            DateTime maxOutTime = DateTime.MaxValue;


                            foreach (CCFTEvent.Event chlEvent in events)
                            {
                                DateTime eventDateTime = chlEvent.OccurrenceTime.AddHours(5);

                                if (chlEvent.EventType == 20001)// In Events
                                {
                                    if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                    {
                                        CardHolderReportInfo reportInfo = cnicWiseReportInfo[cnicNumber + "^" + date.ToString()];
                                        // Returns:
                                        //     One of the following values.Value Description -1 t1 is shorter than t2. 0 t1
                                        //     is equal to t2. 1 t1 is longer than t2.
                                        if (eventDateTime.TimeOfDay < reportInfo.OccurrenceTime.TimeOfDay)
                                        {
                                            // Returns:
                                            //     One of the following values.Value Description -1 t1 is shorter than t2. 0 t1
                                            //     is equal to t2. 1 t1 is longer than t2.
                                            if (TimeSpan.Compare(eventDateTime.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(eventDateTime.TimeOfDay, thEndTime) <= 0)
                                            {
                                                reportInfo.CardNumber = chl.LastName;
                                                reportInfo.OccurrenceTime = eventDateTime;
                                            }
                                            else
                                            {
                                                cnicWiseReportInfo.Remove(cnicNumber + "^" + date.ToString());
                                            }
                                        }


                                    }
                                    else
                                    {
                                        // Returns:
                                        //     One of the following values.Value Description -1 t1 is shorter than t2. 0 t1
                                        //     is equal to t2. 1 t1 is longer than t2.

                                        int startTimeresult = TimeSpan.Compare(eventDateTime.TimeOfDay, thStartTime);//

                                        int endTimeresult = TimeSpan.Compare(eventDateTime.TimeOfDay, thEndTime);//

                                        if (startTimeresult > 0 && endTimeresult <= 0)
                                        {
                                            cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                            {
                                                CardNumber = chl.LastName,
                                                OccurrenceTime = eventDateTime,
                                                FirstName = chl.FirstName,
                                                PNumber = strPnumber,
                                                CNICNumber = cnicNumber,
                                                Department = department,
                                                Section = section,
                                                Cadre = cadre
                                            });
                                        }
                                        else
                                        {
                                            //user checkin before start time range list
                                            if (startTimeresult < 0)
                                            {
                                                if (!checkinBeforeStartTimeLst.Contains(cnicNumber + "^" + date.ToString()))
                                                {
                                                    checkinBeforeStartTimeLst.Add(cnicNumber + "^" + date.ToString());
                                                }
                                            }
                                        }                                        
                                    }
                                }
                            }


                            #endregion
                        }
                    }
                }

                if (cnicWiseReportInfo != null && cnicWiseReportInfo.Keys.Count > 0)
                {
                    for (int i = 0; i < checkinBeforeStartTimeLst.Count; i++)
                    {
                        if (cnicWiseReportInfo.ContainsKey(checkinBeforeStartTimeLst[i])) {
                            cnicWiseReportInfo.Remove(checkinBeforeStartTimeLst[i]);
                        }
                    }


                    foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicWiseReportInfo)
                    {
                        if (reportInfo.Value == null)
                        {
                            continue;
                        }

                        lstCardHolderReportInfo.Add(reportInfo.Value);                       
                    }
                }
         

                #region static data

                //Dictionary<string, Dictionary<string, List<CardHolderInfo>>> data = new Dictionary<string, Dictionary<string, List<CardHolderInfo>>>();
                //Dictionary<string, List<CardHolderInfo>> sections = new Dictionary<string, List<CardHolderInfo>>();

                //  string dep = "Admin";

                //string section = "Accounts";

                //List<CardHolderInfo> cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Qamar",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"                                        
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Abdullah",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "53246",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "D",
                //    FirstName = "Fayyaz",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "6524"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "32014",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Ikram",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3264"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "98765",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Faisal",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "9876"
                //});

                //sections.Add(section, cards);

                //section = "Security";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Mustafa",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Omer",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "98765",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Kamran",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "9876"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "53246",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "D",
                //    FirstName = "Shiraz",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "6524"
                //});

                //sections.Add(section, cards);

                //section = "HR";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Saeed",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Hassan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Ubaid",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "98765",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Haris",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "9876"
                //});


                //sections.Add(section, cards);

                //data.Add(dep, sections);

                //dep = "Quality Assurance";

                //section = "Testers";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Qamar",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Abdullah",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});


                //sections.Add(section, cards);

                //section = "Automation";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Mustafa",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Omer",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "98765",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Kamran",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "9876"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "53246",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "D",
                //    FirstName = "Shiraz",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "6524"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "32014",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Ali",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3264"
                //});


                //sections.Add(section, cards);

                //section = "OutSource Testers";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Saeed",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Hassan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});                
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "53246",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "D",
                //    FirstName = "Abid",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "6524"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "32014",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Mehtab",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3264"
                //});


                //sections.Add(section, cards);

                //data.Add(dep, sections);

                //dep = "Business Analyst";

                //section = "Requirement Gathering";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Qamar",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Abdullah",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});


                //sections.Add(section, cards);

                //section = "Client Dealing";

                //cards = new List<CardHolderInfo>();

                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "12345",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Mustafa",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "1234"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "54321",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "B",
                //    FirstName = "Omer",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "4321"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "Contractor",
                //    CardNumber = "12455",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "C",
                //    FirstName = "Zeeshan",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3214"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "nmpt",
                //    CardNumber = "98765",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Kamran",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "9876"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "53246",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "D",
                //    FirstName = "Shiraz",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "6524"
                //});
                //cards.Add(new CardHolderInfo()
                //{
                //    Cadre = "mpt",
                //    CardNumber = "32014",
                //    CNICNumber = "12345-1234567-1",
                //    Company = "Gallagher",
                //    Crew = "A",
                //    FirstName = "Ali",
                //    OccurrenceTime = DateTime.Now,
                //    PNumber = "3264"
                //});


                //sections.Add(section, cards);


                //data.Add(dep, sections);

               

                #endregion

                if (this.lstCardHolderReportInfo != null && this.lstCardHolderReportInfo.Count > 0)
                {
                    Cursor.Current = currentCursor;
                    this.saveFileDialog1.ShowDialog(this);
                }
                else
                {
                    Cursor.Current = currentCursor;
                    MessageBox.Show(this, "No data exist on current selected date range.");
                }


            }
            catch (Exception exp)
            {
                string exMessage = exp.Message;
                Exception innerException = exp.InnerException;

                while (innerException != null)
                {
                    exMessage = "\n" + innerException.Message;
                    innerException = innerException.InnerException;
                }

                Cursor.Current = currentCursor;
                MessageBox.Show(this, exMessage);
            }

        }


        private bool isValidEntry(string filtersCommaSep, string filterValue)
        {
            bool isValidEntry = true;
            try
            {
                string value = filterValue.ToLower();

                if (filtersCommaSep.ToLower().Contains(value))
                {
                    string[] filters = filtersCommaSep.Split(',');
                    bool skip = true;
                    for (int i = 0; i < filters.Length; i++)
                    {
                        string item = filters[i].Trim();
                        if (value == item.ToLower())
                        {
                            skip = false;
                            break;
                        }
                    }

                    if (skip)
                    {
                        isValidEntry = false;
                    }
                }
                else
                {
                    isValidEntry = false;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
            

            return isValidEntry;
        }


        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.lstCardHolderReportInfo, "Late Arrival Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.lstCardHolderReportInfo, "Late Arrival Report", "Late Arrival Report");
            }
        }

        private void SaveAsPdf(List<CardHolderReportInfo> data, string heading)
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
                                string headerRightText = "Late Time Range: " + this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString();

                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                //pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1000F, 1000F));
                                //  Table table = new Table((new List<float>() { 8F, 100F, 150F, 70F, 250F }).ToArray());
                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(800F, 842F));
                                Table table = new Table((new List<float>() { 50F, 150F, 100F, 100F, 60F, 60F, 60F, 90F }).ToArray());
                                table.SetWidth(720F);
                                table.SetFixedLayout();
                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                //this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data
                                TimeSpan thStartTime = this.dtpLateTimeStart.Value.TimeOfDay;
                                List<string> lstDepartment = new List<string>();
                                List<CardHolderReportInfo> lstDepartmentWiseReportInfo = new List<CardHolderReportInfo>();
                                List<CardHolderReportInfo> lstcardHolderReportInfo = data.OrderBy(o => o.Department).ToList();
                                int lstIndex = lstcardHolderReportInfo.Count - 1;
                                for (int j = 0; j < lstcardHolderReportInfo.Count; j++) 
                                {
                                    CardHolderReportInfo cardHolderReportInfo = lstcardHolderReportInfo[j];
                                    if (lstDepartment.Contains(cardHolderReportInfo.Department))
                                    {
                                        
                                        TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                        string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                        if (timeDiff.Seconds > 30)
                                        {
                                            int minute = timeDiff.Minutes + 1;
                                            diff = timeDiff.Hours + ":" + minute;
                                        }
                                        
                                        cardHolderReportInfo.LateTime = diff;
                                        lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);                                        
                                    }
                                    else
                                    {
                                                                              
                                        lstDepartment.Add(cardHolderReportInfo.Department);

                                        if (lstDepartmentWiseReportInfo.Count > 0)
                                        {
                                            //Department
                                            this.AddDepartmentRow(table, lstDepartmentWiseReportInfo[0].Department);
                                            this.AddTableHeaderRow(table);
                                           
                                            lstDepartmentWiseReportInfo = lstDepartmentWiseReportInfo.OrderByDescending(o => o.LateTime).ToList();
                                            for (int i = 0; i < lstDepartmentWiseReportInfo.Count; i++)
                                            {
                                                CardHolderReportInfo chl = lstDepartmentWiseReportInfo[i];
                                                this.AddTableDataRow(table, chl, i % 2 == 0, i + 1);
                                            }

                                            this.AddNewEmptyRow(table);

                                           
                                            TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                            string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                            if (timeDiff.Seconds > 30)
                                            {
                                                int minute = timeDiff.Minutes + 1;
                                                diff = timeDiff.Hours + ":" + minute;
                                            }
                                            cardHolderReportInfo.LateTime = diff;
                                            lstDepartmentWiseReportInfo = new List<CardHolderReportInfo>();
                                            lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);
                                        }
                                        else
                                        {
                                           
                                            TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                            string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                            if (timeDiff.Seconds > 30)
                                            {
                                                int minute = timeDiff.Minutes + 1;
                                                diff = timeDiff.Hours + ":" + minute;
                                            }
                                            cardHolderReportInfo.LateTime = diff;
                                            lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);
                                        }

                                    }

                                    if (lstIndex == j)
                                    {
                                        if (lstDepartmentWiseReportInfo.Count > 0)
                                        {
                                            //Department
                                            this.AddDepartmentRow(table, lstDepartmentWiseReportInfo[0].Department);
                                            this.AddTableHeaderRow(table);
                                            lstDepartmentWiseReportInfo = lstDepartmentWiseReportInfo.OrderByDescending(o => o.LateTime).ToList();
                                            for (int i = 0; i < lstDepartmentWiseReportInfo.Count; i++)
                                            {
                                                CardHolderReportInfo chl = lstDepartmentWiseReportInfo[i];
                                                this.AddTableDataRow(table, chl, i % 2 == 0, i + 1);
                                            }

                                            this.AddNewEmptyRow(table);
                                        }
                                    }                                                                      
                                }

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

    

        private void SaveAsExcel(List<CardHolderReportInfo> data, string sheetName, string heading)
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

                        work.Column(1).Width = 18.14;
                        work.Column(2).Width = 40;
                        work.Column(3).Width = 25.29;
                        work.Column(4).Width = 18.14;                       
                        work.Column(5).Width = 18.14;
                        work.Column(6).Width = 18.14;
                        work.Column(7).Width = 18.14;
                        work.Column(8).Width = 20.14;


                        //Heading
                        work.Cells["A1:B2"].Merge = true;
                        work.Cells["A1:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["A1:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:B2"].Style.Font.Size = 22;
                        work.Cells["A1:B2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //work.Cells["A1:B2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells["A1:B2"].Value = heading;

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 600);

                        int row = 4;

                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report From: ";
                        work.Cells[row, 2].Value = this.dtpFromDate.Value.ToShortDateString();
                        work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report To:";
                        work.Cells[row, 2].Value = this.dtpToDate.Value.ToShortDateString();
                        work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Late Time Range:";
                        work.Cells[row, 2].Value = this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                        work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;

                        row++;
                        row++;

                        TimeSpan thStartTime = this.dtpLateTimeStart.Value.TimeOfDay;
                        List<string> lstDepartment = new List<string>();
                        List<CardHolderReportInfo> lstDepartmentWiseReportInfo = new List<CardHolderReportInfo>();
                        List<CardHolderReportInfo> lstcardHolderReportInfo = data.OrderBy(o => o.Department).ToList();
                        int lstIndex = lstcardHolderReportInfo.Count - 1;
                        for (int j = 0; j < lstcardHolderReportInfo.Count; j++)
                        {
                            CardHolderReportInfo cardHolderReportInfo = lstcardHolderReportInfo[j];
                            if (lstDepartment.Contains(cardHolderReportInfo.Department))
                            {

                                TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                if (timeDiff.Seconds > 30)
                                {
                                    int minute = timeDiff.Minutes + 1;
                                    diff = timeDiff.Hours + ":" + minute;
                                }

                                cardHolderReportInfo.LateTime = diff;
                                lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);
                            }
                            else
                            {

                                lstDepartment.Add(cardHolderReportInfo.Department);

                                if (lstDepartmentWiseReportInfo.Count > 0)
                                {
                                    

                                        //Department
                                        work.Cells[row, 1].Style.Font.Bold = true;
                                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 1].Value = "Department:";
                                        work.Cells[row, 2].Value = lstDepartmentWiseReportInfo[0].Department;
                                        work.Cells[row, 2].Style.Font.UnderLine = true;
                                        work.Row(row).Height = 20;

                                        row++;

                                        lstDepartmentWiseReportInfo = lstDepartmentWiseReportInfo.OrderByDescending(o => o.LateTime).ToList();
                                        work.Cells[row, 1, row, 8].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                        work.Cells[row, 1, row, 8].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                        work.Cells[row, 1, row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        work.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                                        work.Cells[row, 1, row, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                    work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                    work.Cells[row, 1].Value = "S #.";
                                    work.Cells[row, 2].Value = "First Name";
                                    work.Cells[row, 3].Value = "Section";
                                    work.Cells[row, 4].Value = "Cadre";
                                    work.Cells[row, 5].Value = "P-Number";
                                    work.Cells[row, 6].Value = "Entry Date";
                                    work.Cells[row, 7].Value = "Entry Time";
                                    work.Cells[row, 8].Value = "Total Late (HH:MM)";
                                    work.Row(row).Height = 20;

                                        for (int i = 0; i < lstDepartmentWiseReportInfo.Count; i++)
                                        {
                                            row++;
                                            work.Cells[row, 1, row, 8].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            work.Cells[row, 1, row, 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            work.Cells[row, 1, row, 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            work.Cells[row, 1, row, 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                            work.Cells[row, 1, row, 8].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            work.Cells[row, 1, row, 8].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            work.Cells[row, 1, row, 8].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            work.Cells[row, 1, row, 8].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        if (i % 2 == 0)
                                            {
                                                work.Cells[row, 1, row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                work.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                            }

                                           
                                            work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                            work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                            work.Cells[row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        
                                            CardHolderReportInfo chl = lstDepartmentWiseReportInfo[i];

                                        work.Cells[row, 1].Value = i + 1;
                                        work.Cells[row, 2].Value = chl.FirstName;
                                        work.Cells[row, 3].Value = chl.Section;
                                        work.Cells[row, 4].Value = chl.Cadre;
                                        work.Cells[row, 5].Value = chl.PNumber;
                                        work.Cells[row, 6].Value = chl.OccurrenceTime.ToShortDateString();
                                        work.Cells[row, 7].Value = chl.OccurrenceTime.TimeOfDay.ToString();
                                        work.Cells[row, 8].Value = chl.LateTime;

                                        work.Row(row).Height = 20;
                                        }

                                        row++;
                                        row++;
                                    


                                    TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                    string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                    if (timeDiff.Seconds > 30)
                                    {
                                        int minute = timeDiff.Minutes + 1;
                                        diff = timeDiff.Hours + ":" + minute;
                                    }
                                    cardHolderReportInfo.LateTime = diff;
                                    lstDepartmentWiseReportInfo = new List<CardHolderReportInfo>();
                                    lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);
                                }
                                else
                                {

                                    TimeSpan timeDiff = cardHolderReportInfo.OccurrenceTime.TimeOfDay - thStartTime;
                                    string diff = timeDiff.Hours + ":" + timeDiff.Minutes;
                                    if (timeDiff.Seconds > 30)
                                    {
                                        int minute = timeDiff.Minutes + 1;
                                        diff = timeDiff.Hours + ":" + minute;
                                    }
                                    cardHolderReportInfo.LateTime = diff;
                                    lstDepartmentWiseReportInfo.Add(cardHolderReportInfo);
                                }

                            }

                            if (lstIndex == j)
                            {
                                if (lstDepartmentWiseReportInfo.Count > 0)
                                {
                                    
                                    //Department
                                    work.Cells[row, 1].Style.Font.Bold = true;
                                    work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 1].Value = "Department:";
                                    work.Cells[row, 2].Value = lstDepartmentWiseReportInfo[0].Department;
                                    work.Cells[row, 2].Style.Font.UnderLine = true;
                                    work.Row(row).Height = 20;

                                    row++;

                                    lstDepartmentWiseReportInfo = lstDepartmentWiseReportInfo.OrderByDescending(o => o.LateTime).ToList();
                                    work.Cells[row, 1, row, 8].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    work.Cells[row, 1, row, 8].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 8].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 8].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 8].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                    work.Cells[row, 1, row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    work.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                                    work.Cells[row, 1, row, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                    work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                    work.Cells[row, 1].Value = "S #.";
                                    work.Cells[row, 2].Value = "First Name";
                                    work.Cells[row, 3].Value = "Section";
                                    work.Cells[row, 4].Value = "Cadre";
                                    work.Cells[row, 5].Value = "P-Number";
                                    work.Cells[row, 6].Value = "Entry Date";
                                    work.Cells[row, 7].Value = "Entry Time";
                                    work.Cells[row, 8].Value = "Total Late (HH:MM)";
                                    work.Row(row).Height = 20;

                                    for (int i = 0; i < lstDepartmentWiseReportInfo.Count; i++)
                                    {
                                        row++;
                                        work.Cells[row, 1, row, 8].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                        work.Cells[row, 1, row, 8].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        if (i % 2 == 0)
                                        {
                                            work.Cells[row, 1, row, 8].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            work.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                        }

                                       
                                        work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        


                                        CardHolderReportInfo chl = lstDepartmentWiseReportInfo[i];
                                       
                                        work.Cells[row, 1].Value = i + 1;
                                        work.Cells[row, 2].Value = chl.FirstName;
                                        work.Cells[row, 3].Value = chl.Section;
                                        work.Cells[row, 4].Value = chl.Cadre;
                                        work.Cells[row, 5].Value = chl.PNumber;
                                        work.Cells[row, 6].Value = chl.OccurrenceTime.ToShortDateString();
                                        work.Cells[row, 7].Value = chl.OccurrenceTime.TimeOfDay.ToString();
                                        work.Cells[row, 8].Value = chl.LateTime;

                                        work.Row(row).Height = 20;
                                    }

                                    row++;
                                    row++;
                                }
                            }
                        }



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

        private void AddMainHeading(Table table, string heading)
        {
            Cell headingCell = new Cell(1, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1));
            headingCell.Add(new Paragraph(heading).SetFontSize(20F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
               // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));
            img.SetHeight(40f);
            table.AddCell(headingCell);
           // table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(180F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }

        }

        private void AddDepartmentRow(Table table, string departmentName)
        {
            table.StartNewRow();


            table.AddCell(new Cell().
                  Add(new Paragraph("Department:").
                  SetFontSize(11F).
                  SetBold().
                  SetFontColor(new DeviceRgb(247, 150, 70))).
              SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
              SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
              SetHeight(22F).
              SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                  SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                  SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(departmentName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddSectionRow(Table table, string sectionName)
        {
            table.StartNewRow();

            //table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Section:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(sectionName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddCadreRow(Table table, string cadreName)
        {
            table.StartNewRow();

            //table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Cadre:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(cadreName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddTableHeaderRow(Table table)
        {
            table.StartNewRow();

            table.AddCell(new Cell().
                    Add(new Paragraph("S #.").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                  Add(new Paragraph("First Name").
                  SetFontSize(11F)).
              SetBackgroundColor(new DeviceRgb(253, 233, 217)).
              SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
              SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
              SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                   Add(new Paragraph("Section").
                   SetFontSize(11F)).
               SetBackgroundColor(new DeviceRgb(253, 233, 217)).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Cadre").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                   Add(new Paragraph("P-Number").
                   SetFontSize(11F)).
               SetBackgroundColor(new DeviceRgb(253, 233, 217)).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                   Add(new Paragraph("Entry Date").
                   SetFontSize(11F)).
               SetBackgroundColor(new DeviceRgb(253, 233, 217)).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Entry Time").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                  Add(new Paragraph("Total Late (HH:MM)").
                  SetFontSize(11F)).
              SetBackgroundColor(new DeviceRgb(253, 233, 217)).
              SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
              SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
              SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));


            //table.AddCell(new Cell().
            //        Add(new Paragraph("Crew").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Cadre").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("CNIC Number").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Company Name").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void AddTableDataRow(Table table, CardHolderReportInfo chl, bool altRow,int serNum)
        {
            if (chl == null)
            {
                return;
            }

           

            table.StartNewRow();

            table.AddCell(new Cell().
                   Add(new Paragraph(serNum.ToString()).
                   SetFontSize(11F)).
               SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                   Add(new Paragraph(string.IsNullOrEmpty(chl.Section) ? string.Empty : chl.Section).
                   SetFontSize(11F)).
               SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
               SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
               SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
               SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                Add(new Paragraph(string.IsNullOrEmpty(chl.OccurrenceTime.ToShortDateString()) ? string.Empty : chl.OccurrenceTime.ToShortDateString()).
                SetFontSize(11F)).
            SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                Add(new Paragraph(string.IsNullOrEmpty(chl.OccurrenceTime.TimeOfDay.ToString()) ? string.Empty : chl.OccurrenceTime.TimeOfDay.ToString()).
                SetFontSize(11F)).
            SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                Add(new Paragraph(chl.LateTime).
                SetFontSize(11F)).
            SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Crew) ? string.Empty : chl.Crew).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.CNICNumber) ? string.Empty : chl.CNICNumber).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Company) ? string.Empty : chl.Company).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }


        private void TextBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }

    }

    public class PdfHeaderAndFooter : IEventHandler
    {
        public Document mDoc = null;
        bool mStart = false;
        string mLeftText, mRightText;

        public PdfHeaderAndFooter(Document doc, bool start, string leftText, string rightText)
        {
            this.mDoc = doc;
            this.mStart = start;
            this.mLeftText = leftText;
            this.mRightText = rightText;
        }

        public void HandleEvent(iText.Kernel.Events.Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfCanvas canvas = new PdfCanvas(docEvent.GetPage());
            iText.Kernel.Geom.Rectangle pageSize = docEvent.GetPage().GetPageSize();
            canvas.BeginText();

            try
            {
                canvas.SetFontAndSize(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"), 8);
            }
            catch (IOException e)
            {
            }

            float height = pageSize.GetHeight();
            float width = pageSize.GetWidth();
            float left = pageSize.GetLeft();
            float right = pageSize.GetRight();
            float leftMargin = this.mDoc.GetLeftMargin();
            float rightMargin = this.mDoc.GetRightMargin();
            float topMargin = this.mDoc.GetTopMargin();
            float bottomMargin = this.mDoc.GetBottomMargin();
            float top = pageSize.GetTop();

            if (mStart)
            {
                canvas.SetStrokeColor(new DeviceRgb(247, 150, 70)).MoveTo(leftMargin, height - topMargin+4).LineTo(width - rightMargin, height - topMargin+4).Stroke().SetStrokeColor(iText.Kernel.Colors.Color.BLACK).
                    MoveText(leftMargin, height - topMargin + 14).ShowText(this.mLeftText).
                    MoveText(390, 0).ShowText(this.mRightText).
                   EndText().
                   Release();
            }
            else
            {
                canvas.SetStrokeColor(new DeviceRgb(247, 150, 70)).MoveTo(leftMargin, bottomMargin-2).LineTo(width - rightMargin, bottomMargin-2).Stroke().SetStrokeColor(iText.Kernel.Colors.Color.BLACK).
                    MoveText(leftMargin, bottomMargin - 12).ShowText(this.mLeftText).
                    MoveText(width - rightMargin - 180, 0).ShowText(this.mRightText).
                   EndText().
                   Release();
            }            
        }
    }
}
