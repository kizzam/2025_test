using RPMS.CoreBusiness;
using RPMS.UseCases.PluginInterfaces;


namespace RPMS.UseCases.Pigeons.Interfaces;

public interface IAddPigeonUseCase
{
    static Task<IEnumerable<Pigeon>> IAddPigeonUseCase.ExecuteAsync(string ringno = "");

}