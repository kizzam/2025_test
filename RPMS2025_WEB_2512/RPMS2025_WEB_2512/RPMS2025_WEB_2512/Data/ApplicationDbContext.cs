using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Models;

namespace RPMS2025_WEB_2512.Data;

public class ApplicationDbContext : DbContext
{
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlServer(
            @"Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=RP2025;Integrated Security=True;Encrypt=False");    

    }
    public virtual DbSet<BloodlineModel>? DisplayBloodlinerecords { get; set; }
}

