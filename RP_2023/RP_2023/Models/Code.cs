using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Code
    {
        public string Code1 { get; set; } = null!;
        public string? Description { get; set; }
        public string? Status { get; set; }
        public bool? System { get; set; }
    }
}
