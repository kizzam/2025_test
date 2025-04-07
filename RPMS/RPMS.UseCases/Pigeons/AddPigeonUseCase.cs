using RPMS.CoreBusiness;
using RPMS.UseCases.Pigeons.Interfaces;
using RPMS.UseCases.PluginInterfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//2:12:08

namespace RPMS.UseCases.Pigeons
{
    public class AddPigeonUseCase: IAddPigeonUseCase
 
    {
        private readonly IPigeonRepository pigeonRepository;
            
        public AddPigeonUseCase(IPigeonRepository pigeonRepository)
        {
            this.pigeonRepository = pigeonRepository;
        }

        public async Task ExecuteAsync(Pigeon pigeon)
        {
            if (!await this.pigeonRepository.ExistsAsync(pigeon))
                await this.pigeonRepository.AddPigeonAsync(pigeon);
        }
    }
}
