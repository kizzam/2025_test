using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    public partial class Drug
    {
        public int DrugId { get; set; }
        public string? DrugCode { get; set; }
        public string? DrugDesc { get; set; }
        public string? DrugBrand { get; set; }
        public string? DrugGeneric { get; set; }
        public string? DrugDosageRate { get; set; }
        public string? DrugUse { get; set; }
        public string? DrugSupplier { get; set; }
        public string? DrugManufacturer { get; set; }
        public string? DrugType { get; set; }
        public int? NoteId { get; set; }
        public string? System { get; set; }
        public DateTime? DateCreate { get; set; }
    }
}
