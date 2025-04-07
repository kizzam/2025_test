using System;
using System.Collections.Generic;

namespace BlazorTrainDapper.Models
{
#nullable disable
    [Table("Instrument1")]
    public class Instrument1
    {
        public Instrument1()
        {
            //            Invoices = new HashSet<Invoice>();
        }
        [Key]
        public int InstrumentId { get; set; }

        public string Name { get; set; }
    }
}
