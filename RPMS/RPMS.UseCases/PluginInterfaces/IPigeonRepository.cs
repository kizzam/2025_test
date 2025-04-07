using RPMS.CoreBusiness;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMS.UseCases.PluginInterfaces
{
    public interface IPigeonRepository

    {
        Task AddPigeonAsync(Pigeon pigeon);
        Task<IEnumerable<Pigeon>> GetPigeonsByRingNoAsync(string ringno);

    }
}
