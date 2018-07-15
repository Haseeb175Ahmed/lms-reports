namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Output")]
    public partial class Output
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Output()
        {
            LiftCarLevels = new HashSet<LiftCarLevel>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public bool Pulsed { get; set; }

        public int PulseTime { get; set; }

        public int MaxOnTime { get; set; }

        [Required]
        [StringLength(100)]
        public string OnMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string OffMessage { get; set; }

        [Required]
        [StringLength(100)]
        public string OnStateName { get; set; }

        [Required]
        [StringLength(100)]
        public string OffStateName { get; set; }

        public int Inverted { get; set; }

        [Required]
        [StringLength(100)]
        public string OnTriggerString { get; set; }

        [Required]
        [StringLength(100)]
        public string OffTriggerString { get; set; }

        public int StateMapping { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarLevel> LiftCarLevels { get; set; }
    }
}
