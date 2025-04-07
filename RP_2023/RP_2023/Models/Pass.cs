using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Pass
    {
        public string Username { get; set; } = null!;
        public string? Password { get; set; }
        public int? Accesscode { get; set; }
    }
}
