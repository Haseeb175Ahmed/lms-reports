using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace AttendanceReport
{
    [Table("CategoryInfo")]
    public class CategoryInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int CategoryId { get; set; }

        public string CategoryName { get; set; }

        public string CategoryLocation { get; set; }

        public string CategoryBlockCriteria { get; set; }

    }

    public enum CategoryLocation
    {
        Any = 01,
        Plant =02,
        Colony = 03,          
    }

    public enum CategoryBlockCriteria
    {
        Yes = 001,
        No = 002       
    }
}
