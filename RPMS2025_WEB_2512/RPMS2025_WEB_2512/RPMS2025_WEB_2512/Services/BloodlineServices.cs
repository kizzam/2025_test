using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Components.Pages;
using RPMS2025_WEB_2512.Data;
using Syncfusion.Blazor;
using Syncfusion.Blazor.Data;
using System.Data;
using System.Data.SqlClient;
using System.Reflection.Metadata.Ecma335;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using RPMS2025_WEB_2512.Models;

namespace RPMS2025_WEB_2512.Services
{
    public class BloodlineServices
    {
        protected readonly ApplicationDbContext _dbcontext;

        public BloodlineServices(ApplicationDbContext _db)
        {
            _dbcontext = _db;
        }

        public async Task<BloodlineModel[]> GetbloodlinedetailsAsync()
        {
            return await _dbcontext.DisplayBloodlinerecords.FromSqlRaw("exec spRPMS_Lst_Bloodline_Full").ToArrayAsync();
        }
    }

    public class BloodlineServices_New
    {
        protected readonly ApplicationDbContext _dbcontext;
        private readonly string _connectionString;

        public BloodlineServices_New(ApplicationDbContext _dbContext, string connectionString)
        {
            _dbcontext = _dbContext;
            _connectionString = connectionString;
        }

        public async Task<List<BloodlineModel>> GetBloodlinesAsync()
        {
            string queryString = "SELECT * FROM dbo.BloodLine ORDER BY Code;";
            List<BloodlineModel> bloodlines = new List<BloodlineModel>();

            using (var connection = new SqlConnection(_connectionString))
            {
                SqlDataAdapter adapter = new SqlDataAdapter(queryString, connection);
                DataSet data = new DataSet();
                connection.Open();
                adapter.Fill(data);

                if (data.Tables[0].Rows.Count > 0)
                {
                    bloodlines = data.Tables[0].AsEnumerable().Select(static r => new BloodlineModel
                    {
                        ID = r.Field<int>("ID"),
                        Code = r.Field<string>("Code") ?? string.Empty, // Fix for CS8601
                        Desc = r.Field<string>("Desc"),
                        Status = r.Field<string>("Status"),
                        System = r.Field<bool>("System"),
                        Rating = r.Field<int?>("Rating") ?? 0,
                        RatingBreeding = r.Field<int?>("RatingBreeding") ?? 0, // Fix for CS0019
                        RatingPerf = r.Field<int?>("RatingPerf") ?? 0, // Fix for CS0019
                        RatingCustom = r.Field<int?>("RatingCustom") ?? 0 // Fix for CS0019
                    }).ToList();
                }
            }
            return bloodlines;
        }
    }
}
