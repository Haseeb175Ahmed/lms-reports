namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ArchiveFile")]
    public partial class ArchiveFile
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ArchiveFile()
        {
            ArchiveFileDivisions = new HashSet<ArchiveFileDivision>();
        }

        public int ID { get; set; }

        [Required]
        [StringLength(260)]
        public string Name { get; set; }

        public int Type { get; set; }

        [Required]
        [MaxLength(20)]
        public byte[] Digest { get; set; }

        public DateTime StartTime { get; set; }

        public DateTime EndTime { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ArchiveFileDivision> ArchiveFileDivisions { get; set; }
    }
}
