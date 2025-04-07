using Microsoft.EntityFrameworkCore;
using SalesMgmntApp.Entities;

namespace SalesMgmntApp.Data
{
    public class BirdsMgmntDbContext : DbContext
    {
        public BirdsMgmntDbContext(DbContextOptions<BirdsMgmntDbContext> options) : base(options)

        {

        }

        public DbSet<Employee> Employees { get; set; }
            
        public DbSet<Bird> Birds { get; set; }

        public DbSet<BloodLine> BloodLines { get; set; }

        public DbSet<Colour> Colours { get; set; }


    }
}