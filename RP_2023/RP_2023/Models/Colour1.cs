using System;
using System.Collections.Generic;

namespace RP_2023.Models
{
    [Table("Colour1")]
    public partial class Colour1
    {
        [ExplicitKey] 
        public string Code { get; set; } = null!;
        public string? Desc { get; set; }
        public bool? System { get; set; }
        public DateTime? CreatedDt { get; set; }
        public DateTime? UpdatedDt { get; set; }
    }
}
