using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using RPMS2025_V4.Models;

namespace RPMS2025_V4.Data
{
    public class RPMSdBContext : DbContext
    {
        public RPMSdBContext (DbContextOptions<RPMSdBContext> options)
            : base(options)
        {
        }

        public DbSet<RPMS2025_V4.Models.Bloodline> Bloodline { get; set; } = default!;
    }
}
