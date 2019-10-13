using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport
{
    public class CardHolderReportInfo
    {
        public string CardNumber { get; set; }

        public int TotalDaysCkekin { get; set; }

        public DateTime OccurrenceTime { get; set; }

        public string FirstName { get; set; }

        public string PNumber { get; set; }

        public string Crew { get; set; }

        public string Department { get; set; }

        public string Section { get; set; }

        public string Cadre { get; set; }

        public string CNICNumber { get; set; }

        public string Company { get; set; }

        public int NetNormalHours { get; set; }

        public int OverTimeHours { get; set; }

        public int TotalCallOutHours { get; set; }

        public int NetNormalMinutes { get; set; }

        public int OverTimeMinutes { get; set; }

        public int TotalCallOutMinutes { get; set; }

        public string CallOutFrom { get; set; }

        public string CallOutTo { get; set; }        

        public string ContractorName { get; set; }

        public string Category { get; set; }

        public DateTime BlockedTime { get; set; }

        public DateTime UnBlockedTime { get; set; }

        public string BlockedStatus { get; set; }
        
        public DateTime MinInTime { get; set; }

        public DateTime MaxOutTime { get; set; }

        public DateTime MinCallOutInTime { get; set; }

        public DateTime MaxCallOutOutTime { get; set; }
        public string LateTime { get; set; }

        public string EventType { get; set; }

        public int FtItemId { get; set; }

        public string BlockedBy { get; set; }

        public string BlockedReason { get; set; }

        public string UnBlockedBy { get; set; }

        public string UnBlockedReason { get; set; }

        public string NetNormalTime
        {
            get
            {
                string netNormalTime = string.Empty;

                int netHrs = this.NetNormalHours;
                int netMins = this.NetNormalMinutes;

                while (netMins >= 60)
                {
                    netHrs++;
                    netMins -= 60;
                }

                string netNormalHours = Convert.ToString(netHrs);

                if (netNormalHours.Length < 2)
                {
                    for (int i = netNormalHours.Length; i < 2; i++)
                    {
                        netNormalHours = "0" + netNormalHours;
                    }
                }

                string netNormalMinutes = Convert.ToString(netMins);

                if (netNormalMinutes.Length < 2)
                {
                    for (int i = netNormalMinutes.Length; i < 2; i++)
                    {
                        netNormalMinutes = "0" + netNormalMinutes;
                    }
                }

                netNormalTime = netNormalHours + ":" + netNormalMinutes;

                return netNormalTime;
            }
        }

        public string OverTime
        {
            get
            {
                string overTime = string.Empty;

                int netHrs = this.OverTimeHours;
                int netMins = this.OverTimeMinutes;

                while (netMins >= 60)
                {
                    netHrs++;
                    netMins -= 60;
                }

                string otHours = Convert.ToString(netHrs);

                if (otHours.Length < 2)
                {
                    for (int i = otHours.Length; i < 2; i++)
                    {
                        otHours = "0" + otHours;
                    }
                }

                string otMinutes = Convert.ToString(netMins);

                if (otMinutes.Length < 2)
                {
                    for (int i = otMinutes.Length; i < 2; i++)
                    {
                        otMinutes = "0" + otMinutes;
                    }
                }

                overTime = otHours + ":" + otMinutes;

                return overTime;
            }
        }

        public string NetAndOverTime
        {
            get
            {
                string netAndOverTime = string.Empty;
                int netHrs = this.NetNormalHours + this.OverTimeHours;
                int netMins = this.NetNormalMinutes + this.OverTimeMinutes;

                while(netMins >= 60)
                {
                    netHrs++;
                    netMins -= 60;
                }                 

                string netNormalAndOtHours = Convert.ToString(netHrs);

                if (netNormalAndOtHours.Length < 2)
                {
                    for (int i = netNormalAndOtHours.Length; i < 2; i++)
                    {
                        netNormalAndOtHours = "0" + netNormalAndOtHours;
                    }
                }
                
                string netNormalAndOtMinutes = Convert.ToString(netMins);

                if (netNormalAndOtMinutes.Length < 2)
                {
                    for (int i = netNormalAndOtMinutes.Length; i < 2; i++)
                    {
                        netNormalAndOtMinutes = "0" + netNormalAndOtMinutes;
                    }
                }

                netAndOverTime = netNormalAndOtHours + ":" + netNormalAndOtMinutes;

                return netAndOverTime;
            }
        }

        public string CallOutTime
        {
            get
            {
                string callOutTime = string.Empty;

                int netHrs = this.TotalCallOutHours;
                int netMins = this.TotalCallOutMinutes;

                while (netMins >= 60)
                {
                    netHrs++;
                    netMins -= 60;
                }

                string callOutHours = Convert.ToString(netHrs);

                if (callOutHours.Length < 2)
                {
                    for (int i = callOutHours.Length; i < 2; i++)
                    {
                        callOutHours = "0" + callOutHours;
                    }
                }

                string callOutMinutes = Convert.ToString(netMins);

                if (callOutMinutes.Length < 2)
                {
                    for (int i = callOutMinutes.Length; i < 2; i++)
                    {
                        callOutMinutes = "0" + callOutMinutes;
                    }
                }

                callOutTime = callOutHours + ":" + callOutMinutes;

                return callOutTime;
            }
        }

    }
}
