using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Schedule
    {
        public int Id { get; set; }
        public string? SchCode { get; set; }
        public string? SchDesc { get; set; }
        public string? SchType { get; set; }
        public DateTime? SchDate { get; set; }
        public int? SchSeries { get; set; }
        public int? SchEvent { get; set; }
        public string? SchSystem { get; set; }
        public string? SchStatus { get; set; }
        public string? SchClub { get; set; }
    }
}
