using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class BloodLine
    {
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public string? Status { get; set; }
        public bool? System { get; set; }
        public int? Rating { get; set; }
        public int? RatingBreeding { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingCustom { get; set; }
    }
}
