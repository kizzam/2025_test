﻿using System;
using System.Collections.Generic;

namespace BlazorTrainDapper.Models
{
#nullable disable

    [Table("Customer")]
    public class Customer
    {
        public Customer()
        {
//            Invoices = new HashSet<Invoice>();
        }

        [ExplicitKey]
        public int CustomerId { get; set; } = 0;
        public string FirstName { get; set; } = null!;
        public string LastName { get; set; } = null!;
        public string? Company { get; set; }
        public string? Address { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public string? Country { get; set; }
        public string? PostalCode { get; set; }
        public string? Phone { get; set; }
        public string? Fax { get; set; }
        public string Email { get; set; } = null!;
        public int? SupportRepId { get; set; }

 //       public virtual Employee? SupportRep { get; set; }
 //       public virtual ICollection<Invoice> Invoices { get; set; }
    }
}
