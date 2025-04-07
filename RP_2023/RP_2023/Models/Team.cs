using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Team
    {
        public string Code { get; set; } = null!;
        public string? Description { get; set; }
        public string? Status { get; set; }
    }
}
