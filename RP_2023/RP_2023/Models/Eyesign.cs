using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Eyesign
    {
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public int? Rating { get; set; }
        public string? BestMating { get; set; }
        public bool? System { get; set; }
    }
}
