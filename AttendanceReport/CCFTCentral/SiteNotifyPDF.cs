namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SiteNotifyPDF
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int PDFID { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }
    }
}
