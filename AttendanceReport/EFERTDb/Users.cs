using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Runtime.Serialization;


namespace AttendanceReport
{
    [Table("Users")]
    public class Users
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int UserId { get; set; }

        public string Name { get; set; }

        public string Password { get; set; }

        public string Role { get; set; }

        public DateTime? LastLoginDate { get; set; }

        public String CustomData { get; set; }
    }


    public enum UserRole
    {
        User = 01,
        Admin = 03,
        SiteAdmin = 05
    }

    [Serializable]
    [DataContract]
    public class UserInfoExtended
    {
        [DataMember]
        public UserStatus UserStatus { get; set; }
        [DataMember]
        public DateTime? PasswordChangeDate { get; set; }
        [DataMember]
        public DateTime? UserActiveDate { get; set; }
    }

    public enum UserStatus
    {
        Enabled = 01,
        Disabled = 03
    }
}
