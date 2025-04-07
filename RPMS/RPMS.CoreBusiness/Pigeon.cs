using System.ComponentModel.DataAnnotations;

namespace RPMS.CoreBusiness
{
    public class Pigeon
    {
        [Required]
        public int PigeonId { get; set; }

        [Required]
        public string ringno { get; set; } = string.Empty;

        [Required]
        public string colour { get; set; } = string.Empty;

        public string Sire_id { get; set; } = string.Empty; 

        public string Dam_id { get; set; } = string.Empty; 

    }
}