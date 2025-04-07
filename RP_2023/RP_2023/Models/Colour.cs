using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    [Table("Colour")]
    public partial class Colour
    {
        [ExplicitKey] 
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public bool? System { get; set; }
    }
}
