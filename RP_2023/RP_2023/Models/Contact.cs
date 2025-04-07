using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Contact
    {
        public string Namef { get; set; } = null!;
        public string Namel { get; set; } = null!;
        public string? Address1 { get; set; }
        public string? Address2 { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? Postcode { get; set; }
        public string? Phone { get; set; }
        public string? Fax { get; set; }
        public string? Phonewk { get; set; }
        public string? Phmobil { get; set; }
        public string? Faxwk { get; set; }
        public string? Notes { get; set; }
        public string? Club { get; set; }
        public string? Federation { get; set; }
        public string? Category { get; set; }
    }
}
