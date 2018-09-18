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
        public Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> mData = null;

        public Form1()
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFields(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre, this.cbxCrew);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {

                Cursor.Current = Cursors.WaitCursor;
                this.mData = null;

                DateTime fromDate = this.dtpFromDate.Value.Date;
                DateTime fromDateUtc = fromDate.ToUniversalTime();
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                DateTime toDateUtc = toDate.ToUniversalTime();

                Dictionary<string, CardHolderReportInfo> cnicWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();

                List<string> checkinBeforeStartTimeLst = new List<string>();

                this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();
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

                //List<CCFTEvent.Event> lstEvents = new List<CCFTEvent.Event>() {
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,09,16,07,45,44,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2018,09,16,08,42,44,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //    new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,09,16,09,43,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2018,09,16,10,03,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,09,16,10,15,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2018,09,16,11,58,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2018,09,16,1,15,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2018,09,16,1,58,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    }
                //};

                #endregion



                List<int> inIds = new List<int>();

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
                            if (relatedItem.RelationCode == 0)
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
                                                            where checkin != null && !checkin.CheckedIn && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate &&
                                                                strLstTempCards.Contains(checkin.CardNumber) &&
                                                                (string.IsNullOrEmpty(filterByDepartment) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Department != null &&
                                                                    checkin.CardHolderInfos.Department.DepartmentName.ToLower() == filterByDepartment) ||
                                                                    (checkin.DailyCardHolders != null &&
                                                                    checkin.DailyCardHolders.Department.ToLower() == filterByDepartment))) &&
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
                                                                    checkin.CardHolderInfos.PNumber == filterByPnumber))) &&
                                                                (string.IsNullOrEmpty(filterByCrew) ||
                                                                    ((checkin.CardHolderInfos != null &&
                                                                    checkin.CardHolderInfos.Crew != null &&
                                                                    checkin.CardHolderInfos.Crew.CrewName == filterByCrew)))
                                                            select checkin).ToList();


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

                            //Filter By Department
                            if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() != filterByDepartment.ToLower())
                            {
                                continue;
                            }

                            //Filter By Section
                            if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                            {
                                continue;
                            }

                            //Filter By Cadre
                            if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && section.ToLower() != filterByCadre.ToLower())
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

                            //Filter By Crew
                            if (!string.IsNullOrEmpty(filterByCrew) && crew != filterByCrew)
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


                    this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();

                    foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicWiseReportInfo)
                    {
                        if (reportInfo.Value == null)
                        {
                            continue;
                        }
                        CardHolderReportInfo report = reportInfo.Value;

                        string department = report.Department;
                        string section = report.Section;
                        string cadre = report.Cadre;

                        if (this.mData.ContainsKey(department))
                        {
                            if (this.mData[department].ContainsKey(section))
                            {
                                if (this.mData[department][section].ContainsKey(cadre))
                                {
                                    this.mData[department][section][cadre].Add(report);
                                    this.mData[department][section][cadre].Sort((x, y) => DateTime.Compare(x.OccurrenceTime.Date, y.OccurrenceTime.Date));
                                }
                                else
                                {
                                    this.mData[department][section].Add(cadre, new List<CardHolderReportInfo>() { report });
                                }
                            }
                            else
                            {
                                Dictionary<string, List<CardHolderReportInfo>> cadreWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                                cadreWiseList.Add(cadre, new List<CardHolderReportInfo>() { report });
                                this.mData[department].Add(section, cadreWiseList);
                            }
                        }
                        else
                        {

                            Dictionary<string, List<CardHolderReportInfo>> cadreWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                            cadreWiseList.Add(cadre, new List<CardHolderReportInfo>() { report });

                            Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> sectionWiseReport = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                            sectionWiseReport.Add(section, cadreWiseList);

                            this.mData.Add(department, sectionWiseReport);
                        }
                    }
                }

                #region Old logic
                //Cursor currentCursor = Cursor.Current;
                //try
                //{

                //    Cursor.Current = Cursors.WaitCursor;
                //    this.mData = null;

                //    DateTime fromDate = this.dtpFromDate.Value.Date.ToUniversalTime();
                //    DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59).ToUniversalTime();




                //    CCFTEvent.CCFTEvent ccftEvent = EFERTDbUtility.mCCFTEvent;

                //    CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;
                //    List<int> ids = new List<int>();



                //    #region Events

                //    List<CCFTEvent.Event> lstEvents = (from events in ccftEvent.Events
                //                                       where
                //                                           events != null && events.EventType == 20001 &&
                //                                           events.OccurrenceTime >= fromDate &&
                //                                           events.OccurrenceTime < toDate
                //                                       select events).ToList();



                //    Dictionary<DateTime,Dictionary<int, DateTime>> lstChlEvents = new Dictionary<DateTime, Dictionary<int, DateTime>>();
                //    Dictionary<int, DateTime> dayWiseEvents = null;

                //    //MessageBox.Show("Events Found: " + lstEvents.Count);
                //    foreach (CCFTEvent.Event events in lstEvents)
                //    {
                //        if (events == null || events.RelatedItems == null)
                //        {
                //            continue;
                //        }

                //        foreach (RelatedItem relatedItem in events.RelatedItems)
                //        {
                //            if (relatedItem != null && relatedItem.RelationCode == 0)
                //            {
                //                ids.Add(relatedItem.FTItemID);

                //                if (lstChlEvents.ContainsKey(events.OccurrenceTime.Date))
                //                {
                //                    dayWiseEvents = lstChlEvents[events.OccurrenceTime.Date];

                //                    if (dayWiseEvents.ContainsKey(relatedItem.FTItemID))
                //                    {
                //                        DateTime occurranceTime = dayWiseEvents[relatedItem.FTItemID];

                //                        if (occurranceTime > events.OccurrenceTime)
                //                        {
                //                            dayWiseEvents[relatedItem.FTItemID] = events.OccurrenceTime;
                //                        }

                //                    }
                //                    else
                //                    {
                //                        dayWiseEvents.Add(relatedItem.FTItemID, events.OccurrenceTime);
                //                    }
                //                }
                //                else
                //                {
                //                    dayWiseEvents = new Dictionary<int, DateTime>();
                //                    dayWiseEvents.Add(relatedItem.FTItemID, events.OccurrenceTime);

                //                    lstChlEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                //                }
                //            }
                //        }
                //    }
                //    #endregion

                //    #region CHL
                //    //List<CardholderLocation> cardHolderLocations = (from cardHolder in ccftCentral.CardholderLocations
                //    //                                                where
                //    //                                                   cardHolder != null &&
                //    //                                                   cardHolder.AccessTime >= fromDate &&
                //    //                                                   cardHolder.AccessTime < toDate &&
                //    //                                                   cardHolder.AccessType == 1
                //    //                                                select cardHolder).ToList();


                //    //ids = (from chl in cardHolderLocations
                //    //                 where chl != null
                //    //                 select chl.CardholderID).ToList();
                //    #endregion

                //    //MessageBox.Show("Ids Found: " + lstEvents.Count);

                //    List<Card> cards = (from card in ccftCentral.Cards
                //                        where card != null && ids.Contains(card.CardholderID)
                //                        select card).Distinct().ToList();

                //    #region CHL
                //    //List<Cardholder> cardHolders = (from card in cards
                //    //                                where card != null
                //    //                                select card.Cardholder).ToList();

                //    //Dictionary<int, CardholderLocation> filteredChls =
                //    //                            new Dictionary<int, CardholderLocation>();

                //    //foreach (CardholderLocation chl in cardHolderLocations)
                //    //{
                //    //    if (filteredChls.ContainsKey(chl.CardholderID))
                //    //    {
                //    //        CardholderLocation chlExist = filteredChls[chl.CardholderID];

                //    //        if (chlExist.AccessTime > chl.AccessTime)
                //    //        {
                //    //            filteredChls[chl.CardholderID] = chl;
                //    //        }

                //    //    }
                //    //    else
                //    //    {
                //    //        filteredChls.Add(chl.CardholderID, chl);
                //    //    }
                //    //}
                //    #endregion

                //    this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();
                //    string filterByDepartment = this.cbxDepartments.Text;
                //    string filterBySection = this.cbxSections.Text;
                //    string filerByName = this.tbxName.Text;
                //    string filterByPNumber = this.tbxPNumber.Text;
                //    string filterByCardNumber = this.tbxCarNumber.Text;
                //    string filterByCrew = this.cbxCrew.Text;
                //    string filterByCadre = this.cbxCadre.Text;
                //    string filterByCompany = this.cbxCompany.Text;
                //    string filterByCNIC = this.tbxCnic.Text;

                //    #region CHL
                //    //foreach (KeyValuePair<int, CardholderLocation> chlEvent in filteredChls)
                //    #endregion
                //    #region Events
                //    foreach (KeyValuePair<DateTime, Dictionary<int, DateTime>> chlDateWiseEvent in lstChlEvents)
                //    #endregion
                //    {
                //        foreach (KeyValuePair<int, DateTime> chlEvent in chlDateWiseEvent.Value)
                //        {
                //            if (cards.Exists(c => c.CardholderID == chlEvent.Key))
                //            {
                //                Card card = cards.Find(c => c.CardholderID == chlEvent.Key);
                //                #region CHL
                //                //CardholderLocation cardHolderLocation = chlEvent.Value;
                //                #endregion
                //                #region Events
                //                DateTime occurranceTime = chlEvent.Value.AddHours(5);
                //                #endregion
                //                Cardholder cardHolder = card.Cardholder;

                //                if (cardHolder != null && cardHolder.PersonalDataStrings != null)
                //                {
                //                    string department = (from pds in cardHolder.PersonalDataStrings
                //                                         where pds != null && pds.PersonalDataFieldID == 5043
                //                                         select pds.Value).FirstOrDefault();

                //                    string section = (from pds in cardHolder.PersonalDataStrings
                //                                      where pds != null && pds.PersonalDataFieldID == 12951
                //                                      select pds.Value).FirstOrDefault();

                //                    string crew = (from pds in cardHolder.PersonalDataStrings
                //                                   where pds != null && pds.PersonalDataFieldID == 12869
                //                                   select pds.Value).FirstOrDefault();

                //                    string cadre = (from pds in cardHolder.PersonalDataStrings
                //                                    where pds != null && pds.PersonalDataFieldID == 12952
                //                                    select pds.Value).FirstOrDefault();

                //                    string cnic = (from pds in cardHolder.PersonalDataStrings
                //                                   where pds != null && pds.PersonalDataFieldID == 5051
                //                                   select pds.Value).FirstOrDefault();

                //                    string company = (from pds in cardHolder.PersonalDataStrings
                //                                      where pds != null && pds.PersonalDataFieldID == 5059
                //                                      select pds.Value).FirstOrDefault();

                //                    //if (cadre == "Work Order")
                //                    //{
                //                    //    MessageBox.Show("First Name: "+ cardHolder.FirstName + " cadre: "+ cadre + "Time In: " + occurranceTime.ToString());
                //                    //}

                //                    if (!string.IsNullOrEmpty(department))
                //                    {
                //                        //if (cadre == "Work Order")
                //                        //{
                //                        //    MessageBox.Show("First Name: " + cardHolder.FirstName + " cadre: " + cadre + " Department: " + department + " Time In: " + occurranceTime.ToString());
                //                        //}

                //                        department = department.ToUpper();
                //                        filterByDepartment = filterByDepartment.ToUpper();


                //                        //Filter By Department
                //                        if (!string.IsNullOrEmpty(filterByDepartment) && department != filterByDepartment)
                //                        {
                //                            continue;
                //                        }

                //                        if (!string.IsNullOrEmpty(section))
                //                        {
                //                            //if (cadre == "Work Order")
                //                            //{
                //                            //    MessageBox.Show("First Name: " + cardHolder.FirstName + " cadre: " + cadre + " Department: " + department + " Section: " + section +  " Time In: " + occurranceTime.ToString());
                //                            //}
                //                            section = section.ToUpper();
                //                            filterBySection = filterBySection.ToUpper();


                //                            //Filter By Section
                //                            if (!string.IsNullOrEmpty(filterBySection) && section != filterBySection)
                //                            {
                //                                continue;
                //                            }


                //                            //Filter By Name
                //                            if (!string.IsNullOrEmpty(filerByName) && !cardHolder.FirstName.Contains(filerByName))
                //                            {
                //                                continue;
                //                            }


                //                            //Filter By P-Number
                //                            int? pNumber = cardHolder.PersonalDataIntegers == null || cardHolder.PersonalDataIntegers.Count == 0 ? null : cardHolder.PersonalDataIntegers.ElementAt(0).Value;
                //                            string strPNumber = pNumber == null ? "Nil" : pNumber.ToString();

                //                            if (!string.IsNullOrEmpty(filterByPNumber) && (pNumber == null || pNumber != Convert.ToInt32(filterByPNumber)))
                //                            {
                //                                continue;
                //                            }

                //                            //if (cadre == "Work Order")
                //                            //{
                //                            //    MessageBox.Show("First Name: " + cardHolder.FirstName + " cadre: " + cadre + " P-Number: " + strPNumber +  " Time In: " + occurranceTime.ToString());
                //                            //}

                //                            //Filter By Card Number
                //                            if (!string.IsNullOrEmpty(filterByCardNumber))
                //                            {
                //                                int cardNumber;

                //                                bool parsed = Int32.TryParse(cardHolder.LastName, out cardNumber);

                //                                if (parsed)
                //                                {
                //                                    if (cardNumber != Convert.ToInt32(filterByCardNumber))
                //                                    {
                //                                        continue;
                //                                    }
                //                                }
                //                                else
                //                                {
                //                                    continue;
                //                                }

                //                            }

                //                            //Filter By Crew
                //                            if (!string.IsNullOrEmpty(filterByCrew) && crew != filterByCrew)
                //                            {
                //                                continue;
                //                            }

                //                            //Filter By Cadre
                //                            if (!string.IsNullOrEmpty(filterByCadre) && cadre != filterByCadre)
                //                            {
                //                                continue;
                //                            }

                //                            //Filter By CNIC
                //                            if (!string.IsNullOrEmpty(filterByCNIC) && cnic != filterByCNIC)
                //                            {
                //                                continue;
                //                            }

                //                            //Filter By Company
                //                            if (!string.IsNullOrEmpty(filterByCompany) && company != filterByCompany)
                //                            {
                //                                continue;
                //                            }

                //                            TimeSpan thStartTime = DateTime.MinValue.TimeOfDay;
                //                            TimeSpan thEndTime = DateTime.MaxValue.TimeOfDay;

                //                            if (this.mLateArrivalReport)
                //                            {
                //                                thStartTime = this.dtpLateTimeStart.Value.TimeOfDay;
                //                                thEndTime = this.dtpLateTimeEnd.Value.TimeOfDay;
                //                            }



                //                            #region CHL
                //                            //if (TimeSpan.Compare(cardHolderLocation.AccessTime.AddHours(5).TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(cardHolderLocation.AccessTime.AddHours(5).TimeOfDay, thEndTime) <= 0)
                //                            #endregion
                //                            #region Events
                //                            if (TimeSpan.Compare(occurranceTime.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(occurranceTime.TimeOfDay, thEndTime) <= 0)
                //                            #endregion
                //                            {
                //                                CardHolderReportInfo chi = new CardHolderReportInfo()
                //                                {
                //                                    CardNumber = cardHolder.LastName,
                //                                    FirstName = cardHolder.FirstName,
                //                                    #region CHL
                //                                    //OccurrenceTime = cardHolderLocation.AccessTime.AddHours(5),
                //                                    #endregion
                //                                    #region Events
                //                                    OccurrenceTime = occurranceTime,
                //                                    #endregion
                //                                    PNumber = strPNumber,
                //                                    Crew = crew,
                //                                    Cadre = cadre,
                //                                    Company = company,
                //                                    CNICNumber = cnic
                //                                };


                //                                if (this.mData.ContainsKey(department))
                //                                {
                //                                    if (this.mData[department].ContainsKey(section))
                //                                    {
                //                                        if (this.mData[department][section].ContainsKey(cadre))
                //                                        {
                //                                            List<CardHolderReportInfo> lstchi = this.mData[department][section][cadre];
                //                                            lstchi.Add(chi);
                //                                            this.mData[department][section][cadre] = lstchi;
                //                                        }
                //                                        else
                //                                        {
                //                                            List<CardHolderReportInfo> lstchi = new List<CardHolderReportInfo>();
                //                                            lstchi.Add(chi);
                //                                            this.mData[department][section].Add(cadre, lstchi);
                //                                        }

                //                                    }
                //                                    else
                //                                    {
                //                                        List<CardHolderReportInfo> lstchi = new List<CardHolderReportInfo>();
                //                                        lstchi.Add(chi);
                //                                        Dictionary<string, List<CardHolderReportInfo>> dictCadres = new Dictionary<string, List<CardHolderReportInfo>>();
                //                                        dictCadres.Add(cadre, lstchi);
                //                                        this.mData[department].Add(section, dictCadres);
                //                                    }
                //                                }
                //                                else
                //                                {
                //                                    List<CardHolderReportInfo> lstchi = new List<CardHolderReportInfo>();
                //                                    lstchi.Add(chi);
                //                                    Dictionary<string, List<CardHolderReportInfo>> dictCadres = new Dictionary<string, List<CardHolderReportInfo>>();
                //                                    dictCadres.Add(cadre, lstchi);
                //                                    Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> dictSections = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                //                                    dictSections.Add(section, dictCadres);
                //                                    this.mData.Add(department, dictSections);
                //                                }

                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }                    
                //    }
                #endregion

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

                //this.mData = data;

                #endregion

                if (this.mData != null && this.mData.Count > 0)
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

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.mData, "Late Arrival Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.mData, "Late Arrival Report", "Late Arrival Report");
            }
        }

        private void SaveAsPdf(Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> data, string heading)
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
                                Table table = new Table((new List<float>() { 8F, 100F, 150F, 70F, 250F }).ToArray());


                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                //this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data

                                foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> department in data)
                                {
                                    if (department.Value == null)
                                    {
                                        continue;
                                    }

                                    //Department
                                    this.AddDepartmentRow(table, department.Key);

                                    foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> section in department.Value)
                                    {
                                        if (section.Value == null)
                                        {
                                            continue;
                                        }

                                        //Section
                                        this.AddSectionRow(table, section.Key);


                                        foreach (KeyValuePair<string, List<CardHolderReportInfo>> cadre in section.Value)
                                        {
                                            if (cadre.Value == null)
                                            {
                                                continue;
                                            }

                                            //Cadre
                                            this.AddCadreRow(table, cadre.Key);

                                            //Data
                                            //this.AddNewEmptyRow(table, false);

                                            this.AddTableHeaderRow(table);

                                            for (int i = 0; i < cadre.Value.Count; i++)
                                            {
                                                CardHolderReportInfo chl = cadre.Value[i];
                                                this.AddTableDataRow(table, chl, i % 2 == 0);
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

        private void SaveAsExcel(Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> data, string sheetName, string heading)
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
                        work.Column(2).Width = 25.29;
                        work.Column(3).Width = 15.14;
                        work.Column(4).Width = 54;

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
                        //Sections and Data

                        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> department in data)
                        {
                            if (department.Value == null)
                            {
                                continue;
                            }

                            //Department
                            work.Cells[row, 1].Style.Font.Bold = true;
                            work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            work.Cells[row, 1].Value = "Department:";
                            work.Cells[row, 2].Value = department.Key;
                            work.Cells[row, 2].Style.Font.UnderLine = true;
                            work.Row(row).Height = 20;

                            row++;

                            foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> section in department.Value)
                            {
                                if (section.Value == null)
                                {
                                    continue;
                                }

                                //Section
                                work.Cells[row, 1].Style.Font.Bold = true;
                                work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                work.Cells[row, 1].Value = "Section:";
                                work.Cells[row, 2].Value = section.Key;
                                work.Row(row).Height = 20;

                                //Data
                                row++;

                                foreach (KeyValuePair<string, List<CardHolderReportInfo>> cadre in section.Value)
                                {

                                    if (cadre.Value == null)
                                    {
                                        continue;
                                    }

                                    //Section
                                    work.Cells[row, 1].Style.Font.Bold = true;
                                    work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 1].Value = "Cadre:";
                                    work.Cells[row, 2].Value = cadre.Key;
                                    work.Row(row).Height = 20;

                                    //Data
                                    row++;


                                    work.Cells[row, 1, row, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    work.Cells[row, 1, row, 4].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                    work.Cells[row, 1, row, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    work.Cells[row, 1, row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                                    work.Cells[row, 1, row, 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                    work.Cells[row, 1].Value = "Card Number";
                                    work.Cells[row, 2].Value = "Occurrance Time";
                                    work.Cells[row, 3].Value = "P-Number";
                                    work.Cells[row, 4].Value = "First Name";
                                    work.Row(row).Height = 20;

                                    for (int i = 0; i < cadre.Value.Count; i++)
                                    {
                                        row++;
                                        work.Cells[row, 1, row, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                        work.Cells[row, 1, row, 4].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 4].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 4].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 4].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                        if (i % 2 == 0)
                                        {
                                            work.Cells[row, 1, row, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            work.Cells[row, 1, row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                        }

                                        work.Cells[row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                        work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                        CardHolderReportInfo chl = cadre.Value[i];
                                        work.Cells[row, 1].Value = chl.CardNumber;
                                        work.Cells[row, 2].Value = chl.OccurrenceTime.ToString();
                                        work.Cells[row, 3].Value = chl.PNumber;
                                        work.Cells[row, 4].Value = chl.FirstName;

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
            Cell headingCell = new Cell(2, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
               // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(60F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
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

            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Department:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(departmentName).
                    SetFontSize(11F).SetUnderline()).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddSectionRow(Table table, string sectionName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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

            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Card Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Occurrance Time").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));            
            table.AddCell(new Cell().
                    Add(new Paragraph("P-Number").
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

        private void AddTableDataRow(Table table, CardHolderReportInfo chl, bool altRow)
        {
            if (chl == null)
            {
                return;
            }

            

            table.StartNewRow();

            table.AddCell(new Cell().
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CardNumber) ? string.Empty : chl.CardNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.OccurrenceTime.ToString()) ? string.Empty : chl.OccurrenceTime.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
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
