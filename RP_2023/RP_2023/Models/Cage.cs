using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Cage
    {
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public string? Status { get; set; }
        public bool? System { get; set; }
    }
}
