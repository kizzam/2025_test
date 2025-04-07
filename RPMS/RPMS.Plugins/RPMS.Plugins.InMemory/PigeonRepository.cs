using RPMS.CoreBusiness;    
using RPMS.UseCases.PluginInterfaces;
using System.Security.Cryptography.X509Certificates;

namespace RPMS.Plugins.InMemory
{
    public class PigeonRepository : IPigeonRepository
    {
        private List<Pigeon> _pigeons;

        public PigeonRepository()
        {
            _pigeons = new List<Pigeon>()

            {
               new Pigeon { PigeonId=1, ringno="2022/NRC/850", colour="BB", Sire_id = "2011/GB/12345", Dam_id="2011/sa/12309" },
               new Pigeon { PigeonId=2, ringno="2022/NRC/851", colour="Blue", Sire_id = "2011/NRC/12345", Dam_id="2011/SA/12309" },
               new Pigeon { PigeonId=3, ringno="2022/NRC/852", colour="BB", Sire_id = "2011/QPF/12345", Dam_id="2011/QPF/12309" },
               new Pigeon { PigeonId=4, ringno="2022/NRC/853", colour="White", Sire_id = "2011/GB/12345", Dam_id="2011/SA/12309" },
               new Pigeon { PigeonId=5, ringno="2022/NRC/854", colour="Slate", Sire_id = "2011/GB/12345", Dam_id="2011/QPF/12309" },
               new Pigeon { PigeonId=6, ringno="2022/NRC/855", colour="BBWF", Sire_id = "2011/GB/12345", Dam_id="2011/NRC/12309" },
               new Pigeon { PigeonId=7, ringno="2022/NRC/856", colour="BB", Sire_id = "2011/GB/12345", Dam_id="2011/sa/12309" },
               new Pigeon { PigeonId=8, ringno="2022/NRC/857", colour="BC", Sire_id = "2011/GB/12345", Dam_id="2011/sa/12309" }

            };
        }

        public Task AddPigeonAsync(Pigeon pigeon)
        {
            if (_pigeons.Any(x => x.ringno.Equals(pigeon.ringno, StringComparison.OrdinalIgnoreCase)))
                return Task.CompletedTask;


            var maxid = _pigeons.Max(x => x.PigeonId);
            pigeon.PigeonId = maxid + 1;
            _pigeons.Add(pigeon);

            return Task.CompletedTask;

        }

        public async Task<IEnumerable<Pigeon>> GetPigeonsByRingNoAsync(string ringno)
        {
            if (string.IsNullOrWhiteSpace(ringno)) return await Task.FromResult(_pigeons);

            return _pigeons.Where(x => x.ringno.Contains(ringno, StringComparison.Ordinal));
            
        }

     }
        
}