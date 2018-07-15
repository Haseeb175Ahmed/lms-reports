namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Input")]
    public partial class Input
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Input()
        {
            AlarmKeypads = new HashSet<AlarmKeypad>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public bool Inverted { get; set; }

        [Required]
        [StringLength(100)]
        public string OpenMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string CloseMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string BreakMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string ShortMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string OpenStateName { get; set; }

        [Required]
        [StringLength(100)]
        public string CloseStateName { get; set; }

        [Required]
        [StringLength(100)]
        public string BreakStateName { get; set; }

        [Required]
        [StringLength(100)]
        public string ShortStateName { get; set; }

        public bool NotIsolatable { get; set; }

        public int DiallingXYZ { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmKeypad> AlarmKeypads { get; set; }
    }
}
