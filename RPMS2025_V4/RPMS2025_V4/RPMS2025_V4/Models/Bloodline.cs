using System.ComponentModel.DataAnnotations;

namespace RPMS2025_V4.Models
{
    public class Bloodline
    {
        [Key]
        public int ID { get; set; }
        [Required(ErrorMessage = "Please enter Short Code for the Bloodline 3 Chars Maximum!")]

        [StringLength(3,ErrorMessage = "Please enter Short Code for the Bloodline 3 Chars Maximum!")]
        public string? Code { get; set; }
        [Required]
        [StringLength(30, MinimumLength = 3, ErrorMessage = "Bloodline Name/Desc must be 3 Chars Min - 30 chars Max!")]
        public string? Desc { get; set; }
        [StringLength(1, ErrorMessage = "Status field - Allowed values are 'A' Active or 'I' Inactive")]
        [AllowedValues ("A","I", ErrorMessage = "Status field - Allowed values are 'A' Active or 'I' Inactive")]
        public string? Status { get; set; }
        public bool System { get; set; }
        [Range(0, 9, ErrorMessage = "Please enter a valid Rating in Range 0-9!")]
        public int? Rating { get; set; }
        public int? RatingBreeding { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingCustom { get; set; }
    }
}
