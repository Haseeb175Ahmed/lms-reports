namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ExpiryNotificationRecord")]
    public partial class ExpiryNotificationRecord
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int RecipientID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid ExpiryID { get; set; }

        public int resendByEmail { get; set; }

        public int resendBySMS { get; set; }
    }
}
