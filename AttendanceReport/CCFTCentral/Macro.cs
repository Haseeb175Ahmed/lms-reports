namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Macro")]
    public partial class Macro
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Macro()
        {
            MacroActions = new HashSet<MacroAction>();
            MacroCommandLines = new HashSet<MacroCommandLine>();
            MacroSchedules = new HashSet<MacroSchedule>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(280)]
        public string UserAccount { get; set; }

        [Required]
        [StringLength(620)]
        public string UserPassword { get; set; }

        public bool OperatorConfirmation { get; set; }

        [Required]
        [StringLength(200)]
        public string ConfirmationText { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroAction> MacroActions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroCommandLine> MacroCommandLines { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroSchedule> MacroSchedules { get; set; }
    }
}
