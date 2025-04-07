using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    [Table("Colour2")]
    public partial class Colour2
    {
        [ExplicitKey]
        public int ColourId { get; set; } = 0;
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public bool? System { get; set; }
        public DateTime? CreatedDt { get; set; }
        public DateTime? UpdatedDt { get; set; }
    }
}
