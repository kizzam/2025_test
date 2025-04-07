using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RPMS.CoreBusiness;
using RPMS.UseCases.Pigeons.Interfaces;
using RPMS.UseCases.PluginInterfaces;

namespace RPMS.UseCases.Pigeons
{

    public class ViewPigeonsByRingNoUseCase : IViewPigeonsByRingNoUseCase
    {
        private readonly IPigeonRepository pigeonRepository;

        public ViewPigeonsByRingNoUseCase(IPigeonRepository pigeonRepository)
        {
            this.pigeonRepository = pigeonRepository;
        }


        public async Task<IEnumerable<Pigeon>> ExecuteAsync(string ringno = "")
        {
            return await pigeonRepository.GetPigeonsByRingNoAsync(ringno);

        }
    }
}

//
// Time 1.10.55
//
