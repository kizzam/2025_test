using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Note
    {
        public int NoteId { get; set; }
        public int? EventId { get; set; }
        public int? DrugId { get; set; }
        public DateTime? NoteDate { get; set; }
        public string? NoteMemo { get; set; }
        public string? System { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}
