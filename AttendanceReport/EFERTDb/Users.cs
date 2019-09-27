using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

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
        Normal = 01,
        Admin = 03,
    }
}
