using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.ComponentModel.DataAnnotations;
using RPMS2025_WEB_2512.Models;

namespace RPMS2025_WEB_2512.Data
{
    public class BloodlineData
    {
        string connectionstring = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=RP2025;Integrated Security=True;Encrypt=False";

        public async Task<List<BloodlineModel>> GetBloodlineData()
        {
            using (var context = new BloodLineDbContext(connectionstring))
            {
                var bloodlines = await context.Bloodlines.ToListAsync();
                return bloodlines;

                //return await context.Bloodlines.ToListAsync();
            }
        }
    }
}