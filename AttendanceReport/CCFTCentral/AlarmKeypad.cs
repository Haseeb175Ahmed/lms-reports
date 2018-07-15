namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmKeypad")]
    public partial class AlarmKeypad
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public AlarmKeypad()
        {
            FTItems = new HashSet<FTItem>();
            FTItems1 = new HashSet<FTItem>();
            Inputs = new HashSet<Input>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int ReaderID { get; set; }

        public int PinsMode { get; set; }

        public int InactivityTimeout { get; set; }

        public bool BacklightOn { get; set; }

        public int PinTimeout { get; set; }

        public bool AlarmBeep { get; set; }

        public bool IsPinPad { get; set; }

        public int DisplayAlarmZoneID { get; set; }

        public bool DisableDuringEntry { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Input> Inputs { get; set; }
    }
}
