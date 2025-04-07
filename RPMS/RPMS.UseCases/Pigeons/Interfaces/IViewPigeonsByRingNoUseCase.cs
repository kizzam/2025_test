using RPMS.CoreBusiness;

namespace RPMS.UseCases.Pigeons.Interfaces
{
    public interface IViewPigeonsByRingNoUseCase
    {
        Task<IEnumerable<Pigeon>> ExecuteAsync(string ringno = "");
    }
}