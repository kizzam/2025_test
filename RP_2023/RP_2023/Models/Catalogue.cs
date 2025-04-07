using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Catalogue
    {
        public string CatCode { get; set; } = null!;
        public DateTime? RefDate { get; set; }
        public string? ReportTitle { get; set; }
        public string? ReportTitle2 { get; set; }
        public string? Status { get; set; }
    }
}
