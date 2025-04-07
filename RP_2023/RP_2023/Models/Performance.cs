using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Performance
    {
        public int IdPerfmance { get; set; }
        public int? Yr { get; set; }
        public string? Mark { get; set; }
        public string? RingNo { get; set; }
        public DateTime? PerfDate { get; set; }
        public string? FlyerName { get; set; }
        public string? PntName { get; set; }
        public int? Dist { get; set; }
        public int? _1pos { get; set; }
        public string? _1desc { get; set; }
        public int? _1noFlyers { get; set; }
        public int? _1noBirds { get; set; }
        public int? _2pos { get; set; }
        public string? _2desc { get; set; }
        public int? _2noFlyers { get; set; }
        public int? _2noBirds { get; set; }
        public int? _3pos { get; set; }
        public string? _3desc { get; set; }
        public int? _3noFlyers { get; set; }
        public int? _3noBirds { get; set; }
        public DateTime? DateCreate { get; set; }
        public string? Status { get; set; }
    }
}
