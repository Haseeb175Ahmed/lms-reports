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
    public partial class LMSAttendanceReportForm : Form
    {
        //              Department          Section             Cadre             CNIC Number
        private Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> mData = null;

        public LMSAttendanceReportForm()
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFields(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre, null);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            this.mData = null;
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

            string filterByDepartment = this.cbxDepartments.Text.ToLower();
            string filterBySection = this.cbxSections.Text.ToLower();
            string filerByName = this.tbxName.Text.ToLower();
            string filterByCadre = this.cbxCadre.Text.ToLower();
            string filterByCompany = this.cbxCompany.Text.ToLower();
            string filterByCNIC = this.tbxCnic.Text;
            string filterByPnumber = this.tbxPNumber.Text;

            Dictionary<string, CardHolderReportInfo> cnicDateWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();
            List<string> lstCnics = new List<string>();



            #region Actual Data

            #region Actual Events
            List<CCFTEvent.Event> lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                               where
                                                   events != null && (events.EventType == 20001 || events.EventType == 20003) &&
                                                   events.OccurrenceTime >= fromDateUtc &&
                                                   events.OccurrenceTime < toDateUtc
                                               select events).ToList();

            #endregion

            #region Dummy Events

            //List<CCFTEvent.Event> lstEvents = new List<CCFTEvent.Event>() {
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
            List<int> outIds = new List<int>();

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
                            if (events.EventType == 20001)//In Events
                            {
                                inIds.Add(relatedItem.FTItemID);

                                if (lstChlInEvents.ContainsKey(events.OccurrenceTime.Date))
                                {
                                    if (lstChlInEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                    {
                                        if (!lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Exists(ev => events.OccurrenceTime.TimeOfDay.Hours == ev.OccurrenceTime.TimeOfDay.Hours && events.OccurrenceTime.TimeOfDay.Minutes == ev.OccurrenceTime.TimeOfDay.Minutes))
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

            //MessageBox.Show(this, "In Events Found Keys: " + lstChlInEvents.Keys.Count + " Values: " + lstChlInEvents.Values.Count);
            //MessageBox.Show(this, "Out Events Found Keys: " + lstChlOutEvents.Keys.Count + " Values: " + lstChlOutEvents.Values.Count);

            inCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                             where chl != null && inIds.Contains(chl.FTItemID)
                             select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);

            //MessageBox.Show(this, "In CHls Found Keys: " + inCardHolders.Keys.Count + " Values: " + inCardHolders.Values.Count);

            List<string> strLstTempCards = (from chl in inCardHolders
                                            where chl.Value != null && (chl.Value.FirstName.ToLower().StartsWith("t-") || chl.Value.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-"))
                                            select chl.Value.LastName).ToList();

            //MessageBox.Show(this, "Temp Cards found: " + strLstTempCards.Count);

            List<CheckInAndOutInfo> filteredCheckIns = (from checkin in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                                        where checkin != null && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate &&
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
                                                                checkin.CardHolderInfos.PNumber == filterByPnumber)))
                                                        select checkin).ToList();


            outCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                              where chl != null && outIds.Contains(chl.FTItemID)
                              select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);

            //MessageBox.Show(this, "Filtered Checkins: " + filteredCheckIns.Count);

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
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
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
                                    //}
                                    //else
                                    //{
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
                                    //if (minInTime.TimeOfDay < fdtEndTime)
                                    //{
                                    if (minInTime.TimeOfDay > prevInTime.TimeOfDay)
                                    {
                                        //MessageBox.Show("In Hours set: " + inTime.ToString());
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
                                    //}
                                    //else
                                    //{
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

                            //if (lastCallOutInTimesAfterDayEnd != DateTime.MaxValue)
                            //{
                            //    if (lastCallOutOutTimesAfterDayEnd == DateTime.MaxValue || lastCallOutInTimesAfterDayEnd > lastCallOutOutTimesAfterDayEnd)
                            //    {
                            //        CCFTEvent.Event missingOutEvent = (from events in EFERTDbUtility.mCCFTEvent.Events
                            //                                           where events != null &&
                            //                                                 events.EventType == 20003 &&
                            //                                                 events.RelatedItems != null &&
                            //                                                     (from relatedItem in events.RelatedItems
                            //                                                      where relatedItem != null &&
                            //                                                            relatedItem.RelationCode == 0 &&
                            //                                                            relatedItem.FTItemID == ftItemId
                            //                                                      select relatedItem).Any() &&
                            //                                                 events.OccurrenceTime.Date == date.AddDays(1)
                            //                                           select events).FirstOrDefault();

                            //        if (missingOutEvent == null)
                            //        {
                            //            DateTime callOutDateTime = missingOutEvent.OccurrenceTime.AddHours(0);

                            //            if (date.AddDays(1).DayOfWeek == DayOfWeek.Friday)
                            //            {
                            //                if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                            //                {
                            //                    callOutOutDateTime = callOutDateTime;

                            //                    maxCallOutOutTimeAfterEnd = callOutDateTime;

                            //                    lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                            //                    if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                            //                    }

                            //                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                            //                    }
                            //                }
                            //                else
                            //                {
                            //                    callOutOutDateTime = fdtStartDate;

                            //                    maxCallOutOutTimeAfterEnd = fdtStartDate;

                            //                    lastCallOutOutTimesAfterDayEnd = fdtStartDate;

                            //                    if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = fdtStartDate;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, fdtStartDate);
                            //                    }

                            //                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd[cnicNumber] = fdtStartDate;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, fdtStartDate);
                            //                    }
                            //                }
                            //            }
                            //            else
                            //            {
                            //                if (outDateTime.TimeOfDay < fdtWithBeforeGraceTimeStartTime)
                            //                {
                            //                    callOutOutDateTime = callOutDateTime;

                            //                    maxCallOutOutTimeAfterEnd = callOutDateTime;

                            //                    lastCallOutOutTimesAfterDayEnd = callOutOutDateTime;

                            //                    if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = callOutOutDateTime;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, callOutOutDateTime);
                            //                    }

                            //                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd[cnicNumber] = maxCallOutOutTimeAfterEnd;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, maxCallOutOutTimeAfterEnd);
                            //                    }
                            //                }
                            //                else
                            //                {
                            //                    callOutOutDateTime = ndtStartDate;

                            //                    maxCallOutOutTimeAfterEnd = ndtStartDate;

                            //                    lastCallOutOutTimesAfterDayEnd = ndtStartDate;

                            //                    if (dictLastCallOutOutTimesAfterDayEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd[cnicNumber] = ndtStartDate;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictLastCallOutOutTimesAfterDayEnd.Add(cnicNumber, ndtStartDate);
                            //                    }

                            //                    if (dictCallOutOutTimeAfterEnd.ContainsKey(cnicNumber))
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd[cnicNumber] = ndtStartDate;
                            //                    }
                            //                    else
                            //                    {
                            //                        dictCallOutOutTimeAfterEnd.Add(cnicNumber, ndtStartDate);
                            //                    }
                            //                }
                            //            }
                            //        }
                            //    }
                            //}

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
                                //MessageBox.Show("Lunch Hours: " + lunchHours);
                                //MessageBox.Show("In Hours: " + inTime.TimeOfDay.ToString());
                                //MessageBox.Show("Out Hours: " + outTime.TimeOfDay.ToString());

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

                                //CCFTEvent.Event missingOutEvent = (from events in EFERTDbUtility.mCCFTEvent.Events
                                //                                   where events != null &&
                                //                                         events.EventType == 20003 &&
                                //                                         events.RelatedItems != null &&
                                //                                             (from relatedItem in events.RelatedItems
                                //                                              where relatedItem != null &&
                                //                                                    relatedItem.RelationCode == 0 &&
                                //                                                    relatedItem.FTItemID == ftItemId
                                //                                              select relatedItem).Any() &&
                                //                                         events.OccurrenceTime.Date == date.AddDays(1)
                                //                                   select events).FirstOrDefault();

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

            //MessageBox.Show(this, "Data Collected");
            #endregion

            #region Dummy Data

            //cnicWiseReportInfo.Add("11111-1111111-1^" + DateTime.Now.Date, new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 1",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-1",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-1^" + DateTime.Now.Date.AddDays(1), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(1),
            //    FirstName = "Card Holder 2",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-1",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 6,
            //    OverTimeHours = 0,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.MaxValue,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-1^" + DateTime.Now.Date.AddDays(2), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(2),
            //    FirstName = "Card Holder 3",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-1",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-2^" + DateTime.Now.Date, new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 4",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-2",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-2^" + DateTime.Now.Date.AddDays(1), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(1),
            //    FirstName = "Card Holder 5",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-2",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-2^" + DateTime.Now.Date.AddDays(2), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(2),
            //    FirstName = "Card Holder 6",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-2",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-3^" + DateTime.Now.Date, new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 7",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-3",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-3^" + DateTime.Now.Date.AddDays(1), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(1),
            //    FirstName = "Card Holder 8",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-3",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});

            //cnicWiseReportInfo.Add("11111-1111111-3^" + DateTime.Now.Date.AddDays(2), new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date.AddDays(2),
            //    FirstName = "Card Holder 8",
            //    PNumber = "123456",
            //    CNICNumber = "11111-1111111-3",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2,
            //    MinInTime = DateTime.Now,
            //    MaxOutTime = DateTime.Now
            //});


            #endregion

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

                this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>>();

                foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicDateWiseReportInfo)
                {
                    if (reportInfo.Value == null)
                    {
                        continue;
                    }

                    string cnicNumber = reportInfo.Value.CNICNumber;
                    string department = reportInfo.Value.Department;
                    string section = reportInfo.Value.Section;
                    string cadre = reportInfo.Value.Cadre;

                    if (this.mData.ContainsKey(department))
                    {
                        if (this.mData[department].ContainsKey(section))
                        {
                            if (this.mData[department][section].ContainsKey(cadre))
                            {
                                if (this.mData[department][section][cadre].ContainsKey(cnicNumber))
                                {
                                    this.mData[department][section][cadre][cnicNumber].Add(reportInfo.Value);
                                    this.mData[department][section][cadre][cnicNumber].Sort((x, y) => DateTime.Compare(x.OccurrenceTime.Date, y.OccurrenceTime.Date));
                                }
                                else
                                {
                                    this.mData[department][section][cadre].Add(cnicNumber, new List<CardHolderReportInfo>() { reportInfo.Value });
                                }
                            }
                            else
                            {
                                Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                                cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { reportInfo.Value });
                                this.mData[department][section].Add(cadre, cnicWiseList);
                            }
                        }
                        else
                        {
                            Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                            cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { reportInfo.Value });
                            Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> cadreWiseList = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                            cadreWiseList.Add(cadre, cnicWiseList);
                            this.mData[department].Add(section, cadreWiseList);
                        }
                    }
                    else
                    {
                        Dictionary<string, List<CardHolderReportInfo>> cnicWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                        cnicWiseList.Add(cnicNumber, new List<CardHolderReportInfo>() { reportInfo.Value });

                        Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> cadreWiseList = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                        cadreWiseList.Add(cadre, cnicWiseList);

                        Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> sectionWiseReport = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();
                        sectionWiseReport.Add(section, cadreWiseList);

                        this.mData.Add(department, sectionWiseReport);
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



            //create data object and print report.


        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.mData, "E-Attendance Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.mData, "E-Attendance Report", "E-Attendance Report");
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

                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(980F, 842F));
                                Table table = new Table((new List<float>() { 8F, 90F, 90F, 70F, 70F, 70F, 70F, 90F, 70F, 70F, 70F, 90F }).ToArray());

                                table.SetWidth(900F);
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

                                            this.AddNewEmptyRow(table);
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
                                                    this.AddTableDataRow(table, chl, i % 2 != 0);
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

                        work.Column(1).Width = 20;
                        work.Column(2).Width = 20;
                        work.Column(3).Width = 15.14;
                        work.Column(4).Width = 15.14;
                        work.Column(5).Width = 15.14;
                        work.Column(6).Width = 15.14;
                        work.Column(7).Width = 20;
                        work.Column(8).Width = 15.14;
                        work.Column(9).Width = 15.14;
                        work.Column(10).Width = 15.14;
                        work.Column(11).Width = 20;

                        //Heading
                        work.Cells["A1:C2"].Merge = true;
                        //work.Cells["A1:C2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //work.Cells["A1:C2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:C2"].Style.Font.Size = 22;
                        work.Cells["A1:C2"].Style.Font.Bold = true;
                        work.Cells["A1:C2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:C2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        //work.Cells["A1:B2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells["A1:C2"].Value = heading;

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 1000);

                        int row = 4;

                        work.Cells[row, 1].Style.Font.Bold = true;
                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1, row, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1].Value = "Report From: ";
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 2].Value = this.dtpFromDate.Value.ToShortDateString();
                        work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1, row, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1].Value = "Report To:";
                        work.Cells[row, 1].Style.Font.Bold = true;
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
                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1, row, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;

                        //row++;
                        row++;

                        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>> department in data)
                        {
                            if (department.Value == null)
                            {
                                continue;
                            }

                            //Department
                            work.Cells[row, 1].Style.Font.Bold = true;
                            //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            work.Cells[row, 1, row, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1].Value = "Department:";
                            work.Cells[row, 2].Value = department.Key;
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
                                //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                work.Cells[row, 1, row, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1].Value = "Section:";
                                work.Cells[row, 2].Value = section.Key;
                                work.Row(row).Height = 20;

                                //Data
                                row++;
                                row++;

                                foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> cadre in section.Value)
                                {
                                    if (cadre.Value == null)
                                    {
                                        continue;
                                    }

                                    //Section
                                    work.Cells[row, 1, row, 2].Style.Font.Bold = true;
                                    //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 1, row, 2].Style.Font.UnderLine = true;
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
                                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 1, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 1].Value = "Name:";
                                        work.Cells[row, 2, row, 4].Merge = true;
                                        work.Cells[row, 2, row, 4].Value = chlName;

                                        //Name
                                        work.Cells[row, 5].Style.Font.Bold = true;
                                        //work.Cells[row, 5].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 5, row, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 5, row, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                        work.Cells[row, 5].Value = "P No:";
                                        work.Cells[row, 6, row, 8].Merge = true;
                                        work.Cells[row, 6, row, 8].Value = Convert.ToInt32(pNumber);
                                        work.Row(row).Height = 20;

                                        //Data
                                        row++;


                                        work.Cells[row, 1, row, 11].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        //work.Cells[row, 1, row, 11].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        //work.Cells[row, 1, row, 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                        //work.Cells[row, 1, row, 11].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        //work.Cells[row, 1, row, 11].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        //work.Cells[row, 1, row, 11].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        //work.Cells[row, 1, row, 11].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                        work.Cells[row, 1, row, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        work.Cells[row, 1, row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(191, 191, 191));
                                        work.Cells[row, 1, row, 11].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                        work.Cells[row, 1, row, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                        work.Cells[row, 1].Value = "Date";
                                        work.Cells[row, 2].Value = "Day";
                                        work.Cells[row, 3].Value = "Time In";
                                        work.Cells[row, 4].Value = "Time Out";
                                        work.Cells[row, 5].Value = "Net Hrs";
                                        work.Cells[row, 6].Value = "OT Hrs";
                                        work.Cells[row, 7].Value = "Nrml + OT Hrs";
                                        work.Cells[row, 8].Value = "CO Time In";
                                        work.Cells[row, 9].Value = "CO Time Out";
                                        work.Cells[row, 10].Value = "CO Hrs";
                                        work.Cells[row, 11].Value = "Remarks";

                                        work.Row(row).Height = 20;

                                        for (int i = 0; i < cnicWise.Value.Count; i++)
                                        {
                                            row++;
                                            work.Cells[row, 1, row, 11].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            work.Cells[row, 1, row, 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            work.Cells[row, 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            //work.Cells[row, 1, row, 11].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            //work.Cells[row, 1, row, 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                            //work.Cells[row, 1, row, 11].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            //work.Cells[row, 1, row, 11].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            //work.Cells[row, 1, row, 11].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                            //work.Cells[row, 1, row, 11].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                            if (i % 2 != 0)
                                            {
                                                work.Cells[row, 1, row, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                                work.Cells[row, 1, row, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                                            }

                                            work.Cells[row, 1, row, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            //work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            //work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                            CardHolderReportInfo chl = cnicWise.Value[i];
                                            work.Cells[row, 1].Value = chl.OccurrenceTime.Date.ToShortDateString();
                                            work.Cells[row, 2].Value = chl.OccurrenceTime.Date.DayOfWeek.ToString();

                                            int cellsRemains = 9;

                                            if (chl.MinInTime == DateTime.MaxValue)
                                            {
                                                if (chl.MinCallOutInTime == DateTime.MaxValue)
                                                {
                                                    cellsRemains = 0;

                                                    work.Cells[row, 3, row, 10].Merge = true;
                                                    work.Cells[row, 3, row, 10].Style.Font.Bold = true;
                                                    work.Cells[row, 3, row, 10].Value = "Absent";
                                                }
                                                else
                                                {
                                                    cellsRemains = 4;

                                                    work.Cells[row, 3, row, 7].Merge = true;
                                                    work.Cells[row, 3, row, 7].Style.Font.Bold = true;
                                                    work.Cells[row, 3, row, 7].Value = "Absent";
                                                }
                                            }

                                            if (cellsRemains == 9)
                                            {
                                                work.Cells[row, 3].Value = chl.MinInTime.ToString("HH:mm");
                                                work.Cells[row, 4].Value = chl.MaxOutTime.ToString("HH:mm");
                                                work.Cells[row, 5].Value = chl.NetNormalTime;
                                                work.Cells[row, 6].Value = chl.OverTime;
                                                work.Cells[row, 7].Value = chl.NetAndOverTime;
                                            }

                                            if (cellsRemains == 9 || cellsRemains == 4)
                                            {
                                                work.Cells[row, 8].Value = chl.CallOutFrom;
                                                work.Cells[row, 9].Value = chl.CallOutTo;
                                                work.Cells[row, 10].Value = chl.CallOutTime;
                                            }

                                            work.Cells[row, 11].Value = string.Empty;

                                            work.Row(row).Height = 20;
                                        }

                                        row++;
                                        row++;
                                    }


                                }

                            }
                        }

                        //Sections and Data



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
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            Cell headingCell = new Cell(2, 7);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBold()
                // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
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
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }

        }

        private void AddDepartmentRow(Table table, string departmentName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Department:").
                    SetFontSize(11F).
                    SetBold()).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(departmentName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                    SetBold()).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(sectionName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                    SetBold().SetUnderline()).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F));
            table.AddCell(new Cell().
                    Add(new Paragraph(cadreName).
                    SetFontSize(11F).SetBold().SetUnderline()).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
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

        private void AddChlRow(Table table, string chlName, string pNumber)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Name:").
                    SetFontSize(11F).
                    SetBold()).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F));
            table.AddCell(new Cell(1, 3).
                    Add(new Paragraph(chlName).
                    SetFontSize(11F)).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F));
            //table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                     Add(new Paragraph("P No:").
                     SetFontSize(11F).
                     SetBold()).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                 SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                 SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                 SetHeight(22F));
            table.AddCell(new Cell(1, 2).
                    Add(new Paragraph(pNumber).
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
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
            //        SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

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
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)));
        table.AddCell(new Cell().
                    Add(new Paragraph("Date").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Day").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Time In").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Time Out").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Net Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("OT Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Nrml + OT Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO Time In").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO Time Out").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Remarks").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(191, 191, 191)).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
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
                    Add(new Paragraph(chl.OccurrenceTime.Date.ToShortDateString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            table.AddCell(new Cell().
                    Add(new Paragraph(chl.OccurrenceTime.Date.DayOfWeek.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

            int cellsRemains = 9;

            if (chl.MinInTime == DateTime.MaxValue)
            {
                if (chl.MinCallOutInTime == DateTime.MaxValue)
                {
                    cellsRemains = 0;
                    table.AddCell(new Cell(1, 8).
                            Add(new Paragraph("Absent").
                            SetFontSize(11F).SetBold()).
                        SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                        SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                        SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                        SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
                }
                else
                {
                    cellsRemains = 4;

                    table.AddCell(new Cell(1, 4).
                            Add(new Paragraph("Absent").
                            SetFontSize(11F).SetBold()).
                        SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                        SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                        SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                        SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
                }
            }

            if (cellsRemains == 9)
            {
                table.AddCell(new Cell().
                    Add(new Paragraph(chl.MinInTime.ToString("HH:mm")).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(chl.MaxOutTime.ToString("HH:mm")).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(chl.NetNormalTime).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(chl.OverTime).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(chl.NetAndOverTime).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            }

            if (cellsRemains == 9 || cellsRemains == 4)
            {
                table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CallOutFrom) ? string.Empty : chl.CallOutFrom).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(string.IsNullOrEmpty(chl.CallOutTo) ? string.Empty : chl.CallOutTo).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));

                table.AddCell(new Cell().
                        Add(new Paragraph(chl.CallOutTime).
                        SetFontSize(11F)).
                    SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                    SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            }

            table.AddCell(new Cell().
                    Add(new Paragraph(string.Empty).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(242, 242, 242) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.BLACK, 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }
    }
}
