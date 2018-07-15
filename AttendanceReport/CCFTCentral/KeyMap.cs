namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("KeyMap")]
    public partial class KeyMap
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public KeyMap()
        {
            FTItems = new HashSet<FTItem>();
        }

        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int KeyMapID { get; set; }

        [Required]
        [StringLength(100)]
        public string Name { get; set; }

        public int Key0 { get; set; }

        public int Key1 { get; set; }

        public int Key2 { get; set; }

        public int Key3 { get; set; }

        public int Key4 { get; set; }

        public int Key5 { get; set; }

        public int Key6 { get; set; }

        public int Key7 { get; set; }

        public int Key8 { get; set; }

        public int Key9 { get; set; }

        public int KeyStar { get; set; }

        public int KeyHash { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems { get; set; }
    }
}
