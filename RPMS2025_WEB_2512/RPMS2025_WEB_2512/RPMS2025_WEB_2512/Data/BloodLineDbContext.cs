using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Models;


namespace RPMS2025_WEB_2512.Data

{
    public class BloodLineDbContext : DbContext
    {
        private readonly string _connectionString;

        public BloodLineDbContext(string connectionString)
        {
            _connectionString = connectionString;
        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(_connectionString);
        }

        public DbSet<BloodlineModel> Bloodlines { get; set; }   
    }

}