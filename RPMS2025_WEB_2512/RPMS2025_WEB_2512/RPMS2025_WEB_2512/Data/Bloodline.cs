using Microsoft.EntityFrameworkCore;
using static Grid_CustomAdaptor_EF.Components.Pages.Home;

namespace RPMS2025_WEB_2512.Data
{
    public class Bloodline
    {
        string connectionstring = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=RP2025;Integrated Security=True;Connect Timeout=30;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False";

        public async Task<List<BloodLine>> GetBloodLineAsync()
        {
            using (var context = new BloodlineDbContext(connectionstring)
        }    
    }
   
    public class BloodLineDbContext : DbContext
    {
        private readonly string _connectionstring;

            public BloodLineDbContext(string connectionstring)
        {
            _connectionstring = connectionstring;

        }

        protected overide void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
             optionsBuilder.UseSqlServer(_connectionstring);
        }

        Public dbSet<Bloodlines> Bloodlines { get; set; }
    }
}


