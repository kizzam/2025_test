using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Event
    {
        public int EventId { get; set; }
        public string? EventDesc { get; set; }
        public DateTime? EventDate { get; set; }
        public int? DrugId { get; set; }
        public DateTime? DateFrom { get; set; }
        public DateTime? DateTo { get; set; }
        public string? DosageRate { get; set; }
        public string? Notes { get; set; }
        public int? NoteId { get; set; }
        public string? EventStatus { get; set; }
        public string? EventType { get; set; }
        public string? System { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}
