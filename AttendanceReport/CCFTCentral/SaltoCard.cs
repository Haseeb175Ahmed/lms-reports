namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SaltoCard
    {
        [Key]
        [StringLength(24)]
        public string cardID { get; set; }

        public DateTime shortExpiryDate { get; set; }

        public DateTime? confirmedShortExpiryDate { get; set; }
    }
}
