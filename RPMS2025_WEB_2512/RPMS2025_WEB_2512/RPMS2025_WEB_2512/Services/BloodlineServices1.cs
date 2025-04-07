using Microsoft.EntityFrameworkCore;
using RPMS2025_WEB_2512.Data;
using Syncfusion.Blazor;
using Syncfusion.Blazor.Data;
using System.Reflection.Metadata.Ecma335;
using System.Threading.Tasks;
using RPMS2025_WEB_2512.Components.Pages;

namespace RPMS2025_WEB_2512.Services
{
    public class Bloodlineservices1

    {
        public class CustomAdaptor : DataAdaptor
        {
            public bloodlineData Bloodlineservices = new bloodlineData();

            public override async Task<object> ReadAsync(DataManagerRequest dataManagerRequest, string? key = null);
        {
            IEnumerable<Bloodline> datasource = await Bloodlineservices1.GetBloodlineAsync();
            int totalRecordCount = dataSource.Cast<BloodlineServices1>().Count();
            return dataManagerRequest.RequiresCounts? new DataResult()
            {

                Result = DataSource;
                Count = totalRecordsCount
            } : (object) dataSource;
        }
    }
}
