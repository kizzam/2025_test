using Microsoft.CodeAnalysis.FlowAnalysis.DataFlow;
using Microsoft.EntityFrameworkCore;
using RPMS2025_V4.Data;
using RPMS2025_V4.Models;

namespace RPMS2025_V4.Repositories
{
    public class BloodlineRepository : IBloodlineRepository
    {
        private readonly RPMSdBContext _context;

        public BloodlineRepository(RPMSdBContext context)
        {
            _context = context;
        }

        public async Task Create(Bloodline bloodline)
        {
            _context.Bloodline.Add(bloodline);
            await _context.SaveChangesAsync();
            //throw new NotImplementedException();
        }

        public async Task<List<Bloodline>> Read()
        {
            var result = await _context.Bloodline.ToListAsync();
            return result;

            //throw new NotImplementedException();
        }

        public async Task<Bloodline?> Read(int ID)
        {
            var result = await _context.Bloodline.FirstOrDefaultAsync(e => e.ID == ID);
            return result;

            //throw new NotImplementedException();
        }

        public async Task<bool> Delete(int ID)
        {
            var bloodline = await _context.Bloodline.FirstOrDefaultAsync(e => e.ID == ID);
            if (bloodline is null) return false;

            _context.Bloodline.Remove(bloodline!);
            await _context.SaveChangesAsync();

            return true; // Ensure a boolean value is returned
        }


        public async Task<bool> Update(Bloodline bloodline)
        {
            try
            {
                bool exists = _context.Bloodline.Any(e => e.ID == bloodline.ID);
                if (!exists) return false;
                
                _context.Bloodline.Attach(bloodline).State=EntityState.Modified;
                await _context.SaveChangesAsync();

                return true;
            }
            catch (DbUpdateConcurrencyException)
            {
                throw;
            }
        }
    }
}
