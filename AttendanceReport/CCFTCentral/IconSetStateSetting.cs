namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("IconSetStateSetting")]
    public partial class IconSetStateSetting
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int IconSetID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemTypeID { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int IconState { get; set; }

        public int CoreColour { get; set; }

        [Required]
        public byte[] IconImage { get; set; }

        public Guid GlobalID { get; set; }

        public virtual FTItemTypeState FTItemTypeState { get; set; }

        public virtual IconSet IconSet { get; set; }
    }
}
