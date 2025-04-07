using Microsoft.EntityFrameworkCore;
using RPMS2025_V2.Models;

namespace RPMS2025_V2.Data

{
    public class BloodlineDbContext : DbContext
    {
        private readonly string _connectionString;

        public BloodlineDbContext(string connectionString)
        {
            _connectionString = connectionString;
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(_connectionString);
        }

        //public DbSet<Bloodlines> Bloodlines { get; set; }


        //public DbSet<Bloodlines> GetAllBloodlines()
       // {
            //using (var context = new BloodlineDbContext(_connectionString))
           // {
          //      return context.Bloodlines;
           // }

       // }
    }
}
