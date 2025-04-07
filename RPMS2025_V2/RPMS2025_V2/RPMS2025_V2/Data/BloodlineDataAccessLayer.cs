using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using RPMS2025_V2.Models;
    
    namespace RPMS2025_V2.Data
{
    public class BloodlineDataAccessLayer
    {
        private readonly string _connectionString;

        public BloodlineDataAccessLayer(string connectionString)
        {
           _connectionString = connectionString;
        }

        public void AddBloodline(Bloodlines bloodline)
        {
            using (var context = new BloodlineDbContext(_connectionString))
            {
                context.Bloodlines.Add(bloodline);
                context.SaveChanges();
            }
        }
        public List<Bloodlines> GetBloodlines()
        {
            using (var context = new BloodlineDbContext(_connectionString))
            {
                return context.Bloodlines.ToList();
            }
        }
        public Bloodlines GetBloodline(int id)
        {
            using (var context = new BloodlineDbContext(_connectionString))
            {
#pragma warning disable CS8603 // Possible null reference return.
                return context.Bloodlines.FirstOrDefault(b => b.ID == id);
#pragma warning restore CS8603 // Possible null reference return.
            }
        }
        public void UpdateBloodline(Bloodlines bloodline)
        {
            using (var context = new BloodlineDbContext(_connectionString))
            {
                context.Bloodlines.Update(bloodline);
                context.SaveChanges();
            }
        }
        public async void DeleteBloodline(int id)
        {
            using (var context = new BloodlineDbContext(_connectionString))
            {
                var bloodline = context.Bloodlines.FirstOrDefault(b => b.ID == id);
                if (bloodline != null)
                {
                    context.Bloodlines.Remove(bloodline);
                    await context.SaveChangesAsync();
                }
            }
        }
    }
}
