using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Rating
    {
        public string? CodeRating { get; set; }
        public string? RatingDesc { get; set; }
        public string? Status { get; set; }
        public string Code { get; set; } = null!;
        public bool? System { get; set; }
        public string? Desc { get; set; }
        public int? Level { get; set; }
    }
}
