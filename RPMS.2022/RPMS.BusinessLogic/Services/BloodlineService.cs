using RPMS.DataModels.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMS.BusinessLogic.Services
{
    public class BloodlineService : IBloodlineService
    {
        private readonly RPMSContext? _dbContext = null;
        
        public BloodlineService(RPMSContext dbContext)
        {
            this._dbContext = dbContext;
        }
            
        public List<Bloodline> GetBloodlines()
        {
            return _rpmsContext.Bloodlines.ToList();

        }

        List<BloodLine> IBloodlineService.GetBloodlines()
        {
            throw new NotImplementedException();
        }
    }
}
