using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using RPMS2005_V3.Models;

namespace RPMS2005_V3.Data
{
    public class ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : IdentityDbContext<ApplicationUser>(options)
    {
        public DbSet<Bloodline_V3_Model> Bloodlines_V3 { get; set; }

    }
}
