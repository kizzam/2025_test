// This is the service for the Colour class.
using Dapper;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace RPMS_2023.Data
{
    public class ColourService : IColourService
    {
        // Database connection
        private readonly SqlConnectionConfiguration _configuration;
        public ColourService(SqlConnectionConfiguration configuration)
        {
            _configuration = configuration;
        }
        // Add (create) a Colour table row (SQL Insert)
        // This only works if you're already created the stored procedure.
        public async Task<bool> ColourInsert(Colour colour)
        {
            using (var conn = new SqlConnection(_configuration.Value))
            {
                var parameters = new DynamicParameters();
                parameters.Add("Code", colour.Code, DbType.String);
                parameters.Add("Desc", colour.Desc, DbType.String);
                parameters.Add("System", colour.System, DbType.Boolean);
                // Stored procedure method
                await conn.ExecuteAsync("spColour_Insert", parameters, commandType: CommandType.StoredProcedure);
            }
            return true;
        }
        // Get a list of colour rows (SQL Select)
        // This only works if you're already created the stored procedure.
        public async Task<IEnumerable<Colour>> ColourList()
        {
            IEnumerable<Colour> colours;
            using (var conn = new SqlConnection(_configuration.Value))
            {
                colours = await conn.QueryAsync<Colour>("spColour_List", commandType: CommandType.StoredProcedure);
            }
            return colours;
        }
        // Get one colour based on its ColourID (SQL Select)
        // This only works if you're already created the stored procedure.
        public async Task<Colour> Colour_GetOne(string @Code)
        {
            Colour colour = new Colour();
            var parameters = new DynamicParameters();
            parameters.Add("@Code", Code, DbType.String);
            using (var conn = new SqlConnection(_configuration.Value))
            {
                colour = await conn.QueryFirstOrDefaultAsync<Colour>("spColour_GetOne", parameters, commandType: CommandType.StoredProcedure);
            }
            return colour;
        }
        // Update one Colour row based on its ColourID (SQL Update)
        // This only works if you're already created the stored procedure.
        public async Task<bool> ColourUpdate(Colour colour)
        {
            using (var conn = new SqlConnection(_configuration.Value))
            {
                var parameters = new DynamicParameters();
                parameters.Add("Code", colour.Code, DbType.String);
                parameters.Add("Desc", colour.Desc, DbType.String);
                parameters.Add("System", colour.System, DbType.Boolean);
                await conn.ExecuteAsync("spColour_Update", parameters, commandType: CommandType.StoredProcedure);
            }
            return true;
        }
        // Physically delete one Colour row based on its ColourID (SQL Delete)
        // This only works if you're already created the stored procedure.
        public async Task<bool> ColourDelete(string Code)
        {
            var parameters = new DynamicParameters();
            parameters.Add("@Code", Code, DbType.String);
            using (var conn = new SqlConnection(_configuration.Value))
            {
                await conn.ExecuteAsync("spColour_Delete", parameters, commandType: CommandType.StoredProcedure);
            }
            return true;
        }
    }
}
