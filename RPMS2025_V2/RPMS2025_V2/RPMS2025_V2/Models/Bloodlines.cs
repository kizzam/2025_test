using System.ComponentModel.DataAnnotations;

namespace RPMS2025_V2.Models
{
    public class Bloodlines
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
