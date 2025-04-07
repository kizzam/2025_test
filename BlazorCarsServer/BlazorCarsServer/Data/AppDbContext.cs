using Microsoft.EntityFrameworkCore;
using BlazorCarsServer.Models;


namespace BlazorCarsServer.Data
{

    public class AppDbContext : DbContext
    {
        protected readonly IConfiguration Configuration;

        public AppDbContext(IConfiguration configuration)
        {
            Configuration = configuration;
        }
    
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            options.UseSqlServer(Configuration.GetConnectionString("DbConnectionString"));
        }
        public DbSet<Country> Countries { get; set; }   
    }
}
