namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class DefinedEOLResistanceThreshold
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int EOLSchemeID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int A2DType { get; set; }

        public int TransOCOPH { get; set; }

        public int TransOCOPL { get; set; }

        public int TransOPCLH { get; set; }

        public int TransOPCLL { get; set; }

        public int TransCLSCH { get; set; }

        public int TransCLSCL { get; set; }

        public virtual DefinedEOLResistance DefinedEOLResistance { get; set; }
    }
}
