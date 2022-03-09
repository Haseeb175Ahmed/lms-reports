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
    public partial class ActivityReportForm : Form
    {
        private Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> mData = null;
        public Task<List<string>> mDepartmentTask = null;
        public Task<List<string>> mSectionTask = null;
        public Task<List<string>> mCrewTask = null;
        public Task<List<string>> mCadreTask = null;
        public Task<List<string>> mCompanyTask = null;

        public ActivityReportForm()
        {
            InitializeComponent();
            

            this.mDepartmentTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                List<string> departments = (from pds in ccftCentral.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                            select pds.Value).Distinct().ToList();

                return departments;
            });

            this.mSectionTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                List<string> sections = (from pds in ccftCentral.PersonalDataStrings
                                         where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                         select pds.Value).Distinct().ToList();

                return sections;
            });

            this.mCrewTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                List<string> crews = (from pds in ccftCentral.PersonalDataStrings
                                      where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                      select pds.Value).Distinct().ToList();

                return crews;
            });

            this.mCadreTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                List<string> cadres = (from pds in ccftCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                       select pds.Value).Distinct().ToList();

                return cadres;
            });

            this.mCompanyTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                List<string> companyNames = (from pds in ccftCentral.PersonalDataStrings
                                             where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                             select pds.Value).Distinct().ToList();

                return companyNames;
            });

            this.mDepartmentTask.Start();
            this.mSectionTask.Start();
            this.mCrewTask.Start();
            this.mCadreTask.Start();
            this.mCompanyTask.Start();

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


                this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>>();

                Dictionary<string, List<CardHolderReportInfo>> cnicWiseReportInfo = new Dictionary<string, List<CardHolderReportInfo>>();

                string filterByDepartment = this.cbxDepartments.Text;
                string filterBySection = this.cbxSections.Text;
                string filerByName = this.tbxName.Text;
                string filterByPnumber = this.tbxPNumber.Text;
                string filterByCardNumber = this.tbxCarNumber.Text;
                string filterByCrew = this.cbxCrew.Text;
                string filterByCadre = this.cbxCadre.Text;
                string filterByCompany = this.cbxCompany.Text;
                string filterByCNIC = this.tbxCnic.Text;

                #region Events

                #region Actual Data
                List<CCFTEvent.Event> lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                                   where
                                                       events != null && (events.EventType == 20001 || events.EventType == 20003) &&
                                                       events.OccurrenceTime >= fromDateUtc &&
                                                       events.OccurrenceTime < toDateUtc
                                                   select events).ToList();
                #endregion

                #region Dummy Data

                //    List<CCFTEvent.Event> lstEvents = new List<CCFTEvent.Event>() {
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,05,42,44,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,05,42,44,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //    new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,05,43,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,05,43,20,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,05,58,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,05,58,09,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //    new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,05,59,53,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,05,59,53,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //    new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,06,21,22,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,06,21,22,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,06,49,25,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,06,49,25,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,06,50,19,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,06,50,19,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,08,09,17,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,08,09,17,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,08,09,22,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,08,09,22,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,36,49,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,36,49,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,38,12,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,38,12,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,39,29,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,39,29,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,46,23,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,46,23,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,47,07,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,47,07,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,50,28,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,50,28,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,52,11,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,12,52,11,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,52,51,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,12,52,51,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //     new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,13,58,54,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //      new CCFTEvent.Event() {
                //        EventType = 20001,
                //        OccurrenceTime = new DateTime(2017,09,16,13,58,54,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //       new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,14,00,49,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    },
                //        new CCFTEvent.Event() {
                //        EventType = 20003,
                //        OccurrenceTime = new DateTime(2017,09,16,14,00,49,DateTimeKind.Utc),
                //        RelatedItems = new List<RelatedItem>() {
                //            new RelatedItem() {
                //                RelationCode = 0,
                //                FTItemID = 14716
                //            }
                //        }
                //    }
                //};

                #endregion

                //MessageBox.Show(this, "Events Found:" + lstEvents.Count);
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
                                else if (events.EventType == 20003)//Out events
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
                                    if (dateWiseCheckIn.CheckedIn)
                                    {
                                        cnicWiseReportInfo[cnicNumber + "^" + date.ToString()].Add(new CardHolderReportInfo()
                                        {
                                            OccurrenceTime = dateWiseCheckIn.DateTimeIn,
                                            FirstName = firstName,
                                            PNumber = pNumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre,
                                            EventType = "Enter"
                                        });
                                    }
                                    else
                                    {
                                        cnicWiseReportInfo[cnicNumber + "^" + date.ToString()].Add(new CardHolderReportInfo()
                                        {
                                            OccurrenceTime = dateWiseCheckIn.DateTimeIn,
                                            FirstName = firstName,
                                            PNumber = pNumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre,
                                            EventType = "Enter"
                                        });

                                        cnicWiseReportInfo[cnicNumber + "^" + date.ToString()].Add(new CardHolderReportInfo()
                                        {
                                            OccurrenceTime = dateWiseCheckIn.DateTimeOut,
                                            FirstName = firstName,
                                            PNumber = pNumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre,
                                            EventType = "Exit"
                                        });

                                    }

                                }
                                else
                                {
                                    if (dateWiseCheckIn.CheckedIn)
                                    {
                                        cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new List<CardHolderReportInfo>() { new CardHolderReportInfo()
                                                                    {
                                                                        OccurrenceTime = dateWiseCheckIn.DateTimeIn,
                                                                        FirstName = firstName,
                                                                        PNumber = pNumber,
                                                                        CNICNumber = cnicNumber,
                                                                        Department = department,
                                                                        Section = section,
                                                                        Cadre = cadre,
                                                                        EventType = "Enter"
                                                                    }});
                                    }
                                    else
                                    {
                                        cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new List<CardHolderReportInfo>() { new CardHolderReportInfo()
                                                                    {
                                                                        OccurrenceTime = dateWiseCheckIn.DateTimeIn,
                                                                        FirstName = firstName,
                                                                        PNumber = pNumber,
                                                                        CNICNumber = cnicNumber,
                                                                        Department = department,
                                                                        Section = section,
                                                                        Cadre = cadre,
                                                                        EventType = "Enter"
                                                                    },new CardHolderReportInfo()
                                                                    {
                                                                        OccurrenceTime = dateWiseCheckIn.DateTimeOut,
                                                                        FirstName = firstName,
                                                                        PNumber = pNumber,
                                                                        CNICNumber = cnicNumber,
                                                                        Department = department,
                                                                        Section = section,
                                                                        Cadre = cadre,
                                                                        EventType = "Exit"
                                                                    }});


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
                                DateTime occurranceTime = chlEvent.OccurrenceTime.AddHours(5);

                                if (chlEvent.EventType == 20001)// In Events
                                {
                                    if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                    {
                                        cnicWiseReportInfo[cnicNumber + "^" + date.ToString()].Add(new CardHolderReportInfo()
                                        {
                                            OccurrenceTime = occurranceTime,
                                            FirstName = chl.FirstName,
                                            PNumber = strPnumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre,
                                            EventType = "Enter"
                                        });


                                    }
                                    else
                                    {
                                        cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new List<CardHolderReportInfo>() { new CardHolderReportInfo()
                                                                    {
                                                                        OccurrenceTime = occurranceTime,
                                                                        FirstName = chl.FirstName,
                                                                        PNumber = strPnumber,
                                                                        CNICNumber = cnicNumber,
                                                                        Department = department,
                                                                        Section = section,
                                                                        Cadre = cadre,
                                                                        EventType = "Enter"
                                                                    }});
                                    }
                                }
                                else if (chlEvent.EventType == 20003)// Out Events
                                {
                                    if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                                    {
                                        cnicWiseReportInfo[cnicNumber + "^" + date.ToString()].Add(new CardHolderReportInfo()
                                        {
                                            OccurrenceTime = occurranceTime,
                                            FirstName = chl.FirstName,
                                            PNumber = strPnumber,
                                            CNICNumber = cnicNumber,
                                            Department = department,
                                            Section = section,
                                            Cadre = cadre,
                                            EventType = "Exit"
                                        });


                                    }
                                    else
                                    {
                                        cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new List<CardHolderReportInfo>() { new CardHolderReportInfo()
                                                                    {
                                                                        OccurrenceTime = occurranceTime,
                                                                        FirstName = chl.FirstName,
                                                                        PNumber = strPnumber,
                                                                        CNICNumber = cnicNumber,
                                                                        Department = department,
                                                                        Section = section,
                                                                        Cadre = cadre,
                                                                        EventType = "Exit"
                                                                    }});
                                    }
                                }
                            }
                            
                            
                            #endregion
                        }
                    }
                }

                if (cnicWiseReportInfo != null && cnicWiseReportInfo.Keys.Count > 0)
                {
                    this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>>();

                    foreach (KeyValuePair<string, List<CardHolderReportInfo>> reportInfo in cnicWiseReportInfo)
                    {
                        if (reportInfo.Value == null || reportInfo.Value.Count == 0)
                        {
                            continue;
                        }

                        foreach (CardHolderReportInfo report in reportInfo.Value)
                        {
                            string cnicNumber = report.CNICNumber;
                            string department = report.Department;
                            string section = report.Section;
                            string cadre = report.Cadre;

                            if (this.mData.ContainsKey(department))
                            {
                                if (this.mData[department].ContainsKey(section))
                                {
                                    if (this.mData[department][section].ContainsKey(cadre))
                                    {
                                        if (this.mData[department][section][cadre].ContainsKey(cnicNumber))
                                        {
                                            this.mData[department][section][cadre][cnicNumber].Add(report);
                                            this.mData[department][section][cadre][cnicNumber].Sort((x, y) => DateTime.Compare(x.OccurrenceTime, y.OccurrenceTime));
                                        }
                                        else
                                        {
                                            this.mData[department][section][cadre].Add(cnicNumber, new List<CardHolderReportInfo>() { report });
                                        }
                                    }
                                    else
                                    {
                                        Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                                        cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { report });
                                        this.mData[department][section].Add(cadre, cnicWiseList);
                                    }
                                }
                                else
                                {
                                    Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                                    cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { report });
                                    Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> cadreWiseList = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                                    cadreWiseList.Add(cadre, cnicWiseList);
                                    this.mData[department].Add(section, cadreWiseList);
                                }
                            }
                            else
                            {
                                Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                                cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { report });

                                Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> cadreWiseList = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                                cadreWiseList.Add(cadre, cnicWiseList);

                                Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> sectionWiseReport = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();
                                sectionWiseReport.Add(section, cadreWiseList);

                                this.mData.Add(department, sectionWiseReport);
                            }
                        }

                        
                    }
                }

                if (this.mData != null && this.mData.Count > 0)
                {
                    //Cursor.Current = currentCursor;
                    this.saveFileDialog1.ShowDialog(this);
                }
                else
                {
                    //Cursor.Current = currentCursor;
                    MessageBox.Show(this, "No data exist on current selected date range.");
                }

                #endregion
                

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
                this.SaveAsPdf(this.mData, "Activity Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.mData,  "Activity Report","Activity Report");
            }
        }

        private void SaveAsPdf(Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> data, string heading)
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

                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(750F, 1000F));
                                Table table = new Table((new List<float>() { 8F, 150F, 250F, 100F, 100F  }).ToArray());

                                table.SetWidth(650F);
                                table.SetFixedLayout();

                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data

                                foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> department in data)
                                {
                                    if (department.Value == null)
                                    {
                                        continue;
                                    }

                                    //Department
                                    this.AddDepartmentRow(table, department.Key);

                                    foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> section in department.Value)
                                    {
                                        if (section.Value == null)
                                        {
                                            continue;
                                        }

                                        //Section
                                        this.AddSectionRow(table, section.Key);

                                        foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> cadre in section.Value)
                                        {
                                            if (cadre.Value == null)
                                            {
                                                continue;
                                            }

                                            //Cadre
                                            this.AddCadreRow(table, cadre.Key);


                                            foreach (KeyValuePair<string, List<CardHolderReportInfo>> cnicWise in cadre.Value)
                                            {
                                                if (cnicWise.Value == null || cnicWise.Value.Count == 0)
                                                {
                                                    continue;
                                                }

                                                string chlName = cnicWise.Value[0].FirstName;
                                                string pNumber = cnicWise.Value[0].PNumber;

                                                //cnicWise
                                                this.AddChlRow(table, chlName, pNumber);

                                                //Data
                                                //this.AddNewEmptyRow(table, false);

                                                this.AddTableHeaderRow(table);

                                                for (int i = 0; i < cnicWise.Value.Count; i++)
                                                {
                                                    CardHolderReportInfo chl = cnicWise.Value[i];
                                                    this.AddTableDataRow(table, chl, i % 2 == 0);
                                                }

                                                this.AddNewEmptyRow(table);
                                            }
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

        private void SaveAsExcel(Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> data, string sheetName, string heading)
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

                        work.Column(1).Width = 25.29;
                        work.Column(2).Width = 54;
                        work.Column(3).Width = 15.14;
                        work.Column(4).Width = 18.14;

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

                        //row++;
                        //work.Cells[row, 1].Style.Font.Bold = true;
                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        //work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        //work.Cells[row, 1].Value = "Late Time Range:";
                        //work.Cells[row, 2].Value = this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                        //work.Row(row).Height = 20;

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
                        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> department in data)
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

                            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> section in department.Value)
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

                                foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> cadre in section.Value)
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

                                    foreach (KeyValuePair<string, List<CardHolderReportInfo>> cnicWise in cadre.Value)
                                    {

                                        if (cnicWise.Value == null || cnicWise.Value.Count == 0)
                                        {
                                            continue;
                                        }

                                        string chlName = cnicWise.Value[0].FirstName;
                                        string pNumber = cnicWise.Value[0].PNumber;

                                        //Name
                                        work.Cells[row, 1].Style.Font.Bold = true;
                                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 1].Value = "Name:";
                                        work.Cells[row, 2].Value = chlName;

                                        //Name
                                        work.Cells[row, 3].Style.Font.Bold = true;
                                        work.Cells[row, 3].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 3, row, 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 3, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 3].Value = "P No.:";
                                        work.Cells[row, 4].Value = Convert.ToInt32(pNumber);
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

                                        work.Cells[row, 1].Value = "Occurrance Time";
                                        work.Cells[row, 2].Value = "First Name";
                                        work.Cells[row, 3].Value = "P-Number";
                                        work.Cells[row, 4].Value = "Event Type";

                                        work.Row(row).Height = 20;

                                        for (int i = 0; i < cnicWise.Value.Count; i++)
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

                                            work.Cells[row, 1, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            //work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            //work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                            CardHolderReportInfo chl = cnicWise.Value[i];
                                            work.Cells[row, 1].Value = chl.OccurrenceTime.ToString("d/M/yyyy h:mm:ss tt");
                                            work.Cells[row, 2].Value = string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName;
                                            work.Cells[row, 3].Value = string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber;
                                            work.Cells[row, 4].Value = string.IsNullOrEmpty(chl.EventType) ? string.Empty : chl.EventType;


                                            work.Row(row).Height = 20;
                                        }

                                        row++;
                                        row++;
                                    }


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
            Cell headingCell = new Cell(2, 3);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
                // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(50F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
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

        private void AddChlRow(Table table, string firstName, string pNumber)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Name:").
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
                    Add(new Paragraph(firstName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("P No.:").
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
                    Add(new Paragraph(pNumber).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
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
                    Add(new Paragraph("Occurrance Time").
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
            table.AddCell(new Cell().
                    Add(new Paragraph("P-Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Event Type").
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
                    Add(new Paragraph(chl.OccurrenceTime.ToString("d/M/yyyy h:mm:ss tt")).
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
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.EventType) ? string.Empty : chl.EventType).
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

        private void CbxCrew_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxCrew.Items == null || this.cbxCrew.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> crews = new List<string>() { string.Empty };

                    if (this.mCrewTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                        crews = (from pds in ccftCentral.PersonalDataStrings
                                 where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                 select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        crews.AddRange(this.mCrewTask.Result);
                    }

                    this.cbxCrew.Items.AddRange(crews.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void CbxSections_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxSections.Items == null || this.cbxSections.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> sections = new List<string>() { string.Empty };

                    if (this.mSectionTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                        sections = (from pds in ccftCentral.PersonalDataStrings
                                    where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                    select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        sections.AddRange(this.mSectionTask.Result);
                    }

                    this.cbxSections.Items.AddRange(sections.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void CbxDepartments_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxDepartments.Items == null || this.cbxDepartments.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> departments = new List<string>() { string.Empty };

                    if (this.mDepartmentTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                        departments = (from pds in ccftCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                       select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        departments.AddRange(this.mDepartmentTask.Result);
                    }

                    this.cbxDepartments.Items.AddRange(departments.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void cbxCadre_DropDown(object sender, EventArgs e)
        {
            try
            {
                if (this.cbxCadre.Items == null || this.cbxCadre.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> cadres = new List<string>() { string.Empty };

                    if (this.mCadreTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                        cadres = (from pds in ccftCentral.PersonalDataStrings
                                  where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                  select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        cadres.AddRange(this.mCadreTask.Result);
                    }

                    this.cbxCadre.Items.AddRange(cadres.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void cbxCompany_DropDown(object sender, EventArgs e)
        {
            try
            {
                if (this.cbxCompany.Items == null || this.cbxCompany.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> companyNames = new List<string>() { string.Empty };

                    if (this.mCompanyTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = EFERTDbUtility.mCCFTCentral;

                        companyNames = (from pds in ccftCentral.PersonalDataStrings
                                        where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                        select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        companyNames.AddRange(this.mCompanyTask.Result);
                    }

                    this.cbxCompany.Items.AddRange(companyNames.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }
    }
}
