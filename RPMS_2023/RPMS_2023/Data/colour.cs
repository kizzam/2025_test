using System.ComponentModel.DataAnnotations;
using System;

// This is the model for one row in the database table. You may need to make some adjustments.
namespace RPMS_2023.Data
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
