namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("SessionOperatorClub")]
    public partial class SessionOperatorClub
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SessionID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int OperatorClubID { get; set; }

        public byte State { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual OperatorSession OperatorSession { get; set; }
    }
}
