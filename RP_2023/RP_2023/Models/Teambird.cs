using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Teambird
    {
        public string Code { get; set; } = null!;
        public int Yr { get; set; }
        public string Mark { get; set; } = null!;
        public string Ringno { get; set; } = null!;
        public string? Status { get; set; }
    }
}
