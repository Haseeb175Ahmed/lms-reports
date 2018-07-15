namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ConfigDispatchFieldControllerItemTypePair")]
    public partial class ConfigDispatchFieldControllerItemTypePair
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemTypeID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FieldID { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ControllerItemTypeID { get; set; }

        public int Flag { get; set; }

        public virtual ControllerItemType ControllerItemType { get; set; }

        public virtual FTItemType FTItemType { get; set; }
    }
}
