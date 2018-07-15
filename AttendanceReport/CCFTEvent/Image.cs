namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Image")]
    public partial class Image
    {
        public int ImageID { get; set; }

        public int? SequenceID { get; set; }

        public int RawImageNumber { get; set; }

        public int? ImageNumber { get; set; }

        public int ImageDataLength { get; set; }

        [Column(TypeName = "image")]
        public byte[] ImageData { get; set; }

        public virtual Sequence Sequence { get; set; }
    }
}
