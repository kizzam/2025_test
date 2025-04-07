using System.ComponentModel.DataAnnotations;
using Microsoft.EntityFrameworkCore.Metadata.Internal;

namespace RPMS2025_WEB_2512.Models
{
    public class BloodlineModel
    {
        [Key]
        public int ID { get; set; }
        public required string Code { get; set; }
        public string? Desc { get; set; }
        public string? Status { get; set; }
        public bool? System { get; set; }
        public int? Rating { get; set; }
        public int? RatingBreeding { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingCustom { get; set; }
    }
}
