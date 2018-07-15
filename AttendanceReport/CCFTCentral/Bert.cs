namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Bert")]
    public partial class Bert
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Bert()
        {
            BertDependents = new HashSet<BertDependent>();
            BertDNS = new HashSet<BertDN>();
            InterlockSettings = new HashSet<InterlockSetting>();
            FTItems = new HashSet<FTItem>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? SerialNumber { get; set; }

        public int? IPAddress { get; set; }

        public int? IPPort { get; set; }

        [StringLength(256)]
        public string RASEntryName { get; set; }

        [StringLength(50)]
        public string RASUserName { get; set; }

        [StringLength(50)]
        public string RASPassword { get; set; }

        public int? IPGateway { get; set; }

        public int? IPSubnet { get; set; }

        [StringLength(50)]
        public string Timezone { get; set; }

        [StringLength(50)]
        public string ServerPhoneNumber { get; set; }

        public int? ServerIPAddressOut { get; set; }

        public int? ServerIPAddressIn { get; set; }

        public int? BertServerIPAddress { get; set; }

        public int? BertServerIPAddressConfig { get; set; }

        [StringLength(80)]
        public string CertType { get; set; }

        [Required]
        [StringLength(1024)]
        public string CurrentDLLs { get; set; }

        public int SpecialCalendar { get; set; }

        public bool HasSecondGbus { get; set; }

        public int OfflineTimeout { get; set; }

        public bool UseDNS { get; set; }

        [Required]
        [StringLength(255)]
        public string Hostname { get; set; }

        public int ReaderVariant { get; set; }

        public int EnableSecondaryLink { get; set; }

        public int UseSiteUSBDefaults { get; set; }

        public int AllowUSBUpgrades { get; set; }

        public int DownloadLimit { get; set; }

        public bool CommunicationEnabled { get; set; }

        public bool USBRequiredSoftware { get; set; }

        public bool USBRequiredCustom { get; set; }

        public float LowPowerThreshold { get; set; }

        public bool GbusIsAperio { get; set; }

        public bool Gbus2IsAperio { get; set; }

        public int Bus1Type { get; set; }

        public int Bus2Type { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<BertDependent> BertDependents { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<BertDN> BertDNS { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<InterlockSetting> InterlockSettings { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems { get; set; }
    }
}
