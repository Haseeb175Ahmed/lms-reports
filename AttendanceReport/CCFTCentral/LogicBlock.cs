namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LogicBlock")]
    public partial class LogicBlock
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public LogicBlock()
        {
            LogicBlockRelations = new HashSet<LogicBlockRelation>();
            LogicBlockRules = new HashSet<LogicBlockRule>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int UnknownRule { get; set; }

        public int TimingMode { get; set; }

        public int TimingMaximumOn { get; set; }

        public int TimingPulseDuration { get; set; }

        public int TimingPulseOffTime { get; set; }

        public int TimingDelayedOff { get; set; }

        public int TimingDelayedOn { get; set; }

        public int? TimingClearedByInputID { get; set; }

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

        public int StateMapping { get; set; }

        public int DiallingXYZ { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LogicBlockRelation> LogicBlockRelations { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LogicBlockRule> LogicBlockRules { get; set; }
    }
}
