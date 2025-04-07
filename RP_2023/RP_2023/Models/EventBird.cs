using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class EventBird
    {
        public int IdEventBird { get; set; }
        public int? EventId { get; set; }
        public int Yr { get; set; }
        public string Mark { get; set; } = null!;
        public string Ringno { get; set; } = null!;
        public string? Notes { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}
