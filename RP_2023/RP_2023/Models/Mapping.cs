using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Mapping
    {
        public int IdMapping { get; set; }
        public string MapFieldFrom { get; set; } = null!;
        public string MapFieldTo { get; set; } = null!;
        public string? MapType { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}
