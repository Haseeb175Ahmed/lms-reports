namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OnlineArchiveFile")]
    public partial class OnlineArchiveFile
    {
        public DateTime StartTime { get; set; }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FileID { get; set; }

        public DateTime EndTime { get; set; }

        public byte RemovingEvents { get; set; }

        public byte RemovingImages { get; set; }
    }
}
