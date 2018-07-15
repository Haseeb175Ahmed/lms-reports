namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AlarmListNavPanelDisplayColumn")]
    public partial class AlarmListNavPanelDisplayColumn
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public Guid GlobalID { get; set; }

        [Required]
        [StringLength(200)]
        public string ColumnTitle { get; set; }

        public int DefaultWidth { get; set; }

        public int DisplayOrder { get; set; }

        public bool Filterable { get; set; }

        [Key]
        [Column(Order = 1)]
        [StringLength(200)]
        public string FieldBrowseName { get; set; }

        public virtual Viewer Viewer { get; set; }
    }
}
