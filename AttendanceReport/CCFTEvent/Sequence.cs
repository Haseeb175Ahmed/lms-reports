namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Sequence")]
    public partial class Sequence
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Sequence()
        {
            Images = new HashSet<Image>();
            TriggerSequencePairs = new HashSet<TriggerSequencePair>();
        }

        public int SequenceID { get; set; }

        public int CameraID { get; set; }

        public int CameraRotation { get; set; }

        public int MinFix { get; set; }

        public int? ArchiveFileID { get; set; }

        public int DivisionID { get; set; }

        [Required]
        [StringLength(110)]
        public string Name { get; set; }

        [Required]
        [StringLength(200)]
        public string Description { get; set; }

        public int? DVRTypeID { get; set; }

        public virtual DVRData DVRData { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Image> Images { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TriggerSequencePair> TriggerSequencePairs { get; set; }
    }
}
