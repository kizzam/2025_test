using RPMS.DataModels.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPMS.BusinessLogic.Services
{
    public interface IBloodlineService
    {
        List<BloodLine> GetBloodlines();
    }
}
