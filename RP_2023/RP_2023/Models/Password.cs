using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Password
    {
        public string User { get; set; } = null!;
        public string? Pword { get; set; }
        public string? Level { get; set; }
        public bool? Butt1 { get; set; }
        public bool? Butt2 { get; set; }
        public bool? Butt3 { get; set; }
        public bool? Butt4 { get; set; }
        public bool? Butt5 { get; set; }
        public bool? Butt6 { get; set; }
        public bool? Butt7 { get; set; }
        public string? CreatedBy { get; set; }
        public string? LastChangedBy { get; set; }
        public int? DateLastChange { get; set; }
        public string? Pword1 { get; set; }
        public string? Pword2 { get; set; }
        public string? Pword3 { get; set; }
        public DateTime? PwexpiryDte { get; set; }
    }
}
