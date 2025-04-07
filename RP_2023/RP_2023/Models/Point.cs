using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Point
    {
        public string CodePnt { get; set; } = null!;
        public string? Pntname { get; set; }
        public string? State { get; set; }
        public int? LatDeg { get; set; }
        public int? LatMin { get; set; }
        public double? LatSec { get; set; }
        public int? LngDeg { get; set; }
        public int? LngMin { get; set; }
        public double? LngSec { get; set; }
        public string? Series { get; set; }
        public double? DistSpec { get; set; }
        public double? DistCalcGcmTdpf { get; set; }
        public double? DistCalcGcmUk { get; set; }
        public double? DistCalcGcmIntl { get; set; }
        public double? DistCalcGeod { get; set; }
        public double? DistGen { get; set; }
        public string? Notes { get; set; }
        public bool? System { get; set; }
        public string? MyPoint { get; set; }
        public string? Club { get; set; }
        public double? DistCalcGcmTdpfyds { get; set; }
        public double? DistCalcGcmUkyds { get; set; }
        public double? DistCalcGcmIntlyds { get; set; }
        public double? DistCalcGeodyds { get; set; }
        public string? DistType { get; set; }
    }
}
