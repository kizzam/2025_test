using System.ComponentModel.DataAnnotations;

namespace RPMS2005_V3.Models
{
    public class Bloodline_V3_Model
    {
        [Key]
        public int ID { get; set; }
        [Required]
        public required string Code { get; set; }
        public required string Desc { get; set; }
        public required string Status { get; set; }
        public bool? System { get; set; }
        public int? Rating { get; set; }
        public int? RatingBreeding { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingCustom { get; set; }
    }

}
