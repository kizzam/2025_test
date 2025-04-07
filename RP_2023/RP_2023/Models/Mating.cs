using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Mating
    {
        public int Id { get; set; }
        public string? Code { get; set; }
        public int? Rating { get; set; }
        public int? RatingPerf { get; set; }
        public int? RatingBloodLine { get; set; }
        public int? RatingCustom { get; set; }
        public string? StsSystem { get; set; }
        public string? StsMating { get; set; }
        public string? StsCurrent { get; set; }
        public string? StsCustom { get; set; }
        public int? SireYr { get; set; }
        public string? SireMark { get; set; }
        public string? SireRingNo { get; set; }
        public int? DamYr { get; set; }
        public string? DamMark { get; set; }
        public string? DamRingNo { get; set; }
        public string? Notes { get; set; }
    }
}
