using System.ComponentModel.DataAnnotations;

namespace BlazorCRUDsf.Data
{
    public class Colour
    {
        [Required]
        [StringLength(5)]
        public string? Code { get; set; }
        [StringLength(16)]
        public string? Desc { get; set; }
        public bool System { get; set; }
    }
}
