using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using RPMS_2023.Data;
using System.Reflection.Metadata.Ecma335;
using Dapper;

namespace RPMS_2023.Controllers
{
    [Route("api/[controller]")]
    [ApiController]

    public class RPMS2023Controller : ControllerBase

    {
        private readonly IConfiguration _config;

        public RPMS2023Controller(IConfiguration config)
        {
            _config = config;
        }

        [HttpGet]
        public async Task<ActionResult<List<Colour>>> ColourList()
        {
            using var connection = new SqlConnection( _config.GetConnectionString("DefaultConnection"));
            var colour
                = await Connection.QueryAsync<Colour>("select * from colour");
        Return OK(colour);
        }
    }
}
