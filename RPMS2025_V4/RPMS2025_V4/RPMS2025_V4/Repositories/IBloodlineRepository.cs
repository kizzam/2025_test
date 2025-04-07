using RPMS2025_V4.Models;

namespace RPMS2025_V4.Repositories
{
    public interface IBloodlineRepository
    {
        Task Create(Bloodline bloodline);
        
        Task<List<Bloodline>> Read();

        Task<Bloodline?> Read(int ID);

        Task <bool> Update(Bloodline bloodline);

        Task<bool> Delete(int ID);
    }
}
