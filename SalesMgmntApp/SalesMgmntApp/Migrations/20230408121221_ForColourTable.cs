using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SalesMgmntApp.Migrations
{
    /// <inheritdoc />
    public partial class ForColourTable : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Colours",
                columns: table => new
                {
                    ColourId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    ColourCode = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    ColourFull = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    ColourDescription = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false),
                    DTModif = table.Column<DateTime>(name: "DT_Modif", type: "datetime2", nullable: false),
                    DTCreate = table.Column<DateTime>(name: "DT_Create", type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Colours", x => x.ColourId);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Colours");
        }
    }
}
