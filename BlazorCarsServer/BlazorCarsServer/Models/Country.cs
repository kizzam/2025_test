using System.ComponentModel.DataAnnotations;

namespace BlazorCarsServer.Models
{
    public class Country
    {

        public int Id { get; set; }

        [Required]
        [StringLength(10, ErrorMessage = "Name is too long.")]
        public string? Name { get; set; }

        [Required]
        [StringLength(5, ErrorMessage = "MUST Provide Population")]
        public string? Population { get; set; }

        public bool HaveMountains { get; set; }
    }
}