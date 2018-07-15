namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmInstruction")]
    public partial class AlarmInstruction
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public AlarmInstruction()
        {
            EventGroupDefaultInstructions = new HashSet<EventGroupDefaultInstruction>();
            EventGroupItemAlarmInstructions = new HashSet<EventGroupItemAlarmInstruction>();
            EventTypeDefaultInstructions = new HashSet<EventTypeDefaultInstruction>();
            EventTypeItemAlarmInstructions = new HashSet<EventTypeItemAlarmInstruction>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Column(TypeName = "text")]
        public string Message { get; set; }

        public int? SoundID { get; set; }

        public int? PictureID { get; set; }

        public int SoundAutoPlay { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual Picture Picture { get; set; }

        public virtual Sound Sound { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupDefaultInstruction> EventGroupDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroupItemAlarmInstruction> EventGroupItemAlarmInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeDefaultInstruction> EventTypeDefaultInstructions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventTypeItemAlarmInstruction> EventTypeItemAlarmInstructions { get; set; }
    }
}
