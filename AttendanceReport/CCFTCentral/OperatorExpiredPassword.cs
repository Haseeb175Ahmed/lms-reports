namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class OperatorExpiredPassword
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int OperatorID { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(50)]
        public string LogonName { get; set; }

        [Key]
        [Column(Order = 2)]
        public DateTime EntryDate { get; set; }

        [Key]
        [Column(Order = 3)]
        [MaxLength(100)]
        public byte[] EncryptedPassword { get; set; }

        public virtual Cardholder Cardholder { get; set; }
    }
}
