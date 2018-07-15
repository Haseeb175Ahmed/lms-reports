namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Operator")]
    public partial class Operator
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Operator()
        {
            OperatorSessions = new HashSet<OperatorSession>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(50)]
        public string LogonName { get; set; }

        public DateTime LastPasswordChange { get; set; }

        public bool IsSystemOperator { get; set; }

        [Required]
        [MaxLength(100)]
        public byte[] EncryptedPassword { get; set; }

        public bool AllowDropOn { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorSession> OperatorSessions { get; set; }
    }
}
