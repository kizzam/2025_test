using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace SalesMgmntApp.Migrations
{
    /// <inheritdoc />
    public partial class initialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Birds",
                columns: table => new
                {
                    BirdId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    RingYr = table.Column<int>(type: "int", nullable: false),
                    RingMark = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    RingNo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    BirdName = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DOB = table.Column<DateTime>(type: "datetime2", nullable: false),
                    Gender = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Breeder = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    BredFor = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    BloodLineId = table.Column<int>(type: "int", nullable: false),
                    ColourId = table.Column<int>(type: "int", nullable: false),
                    EyeCode = table.Column<string>(name: "Eye_Code", type: "nvarchar(max)", nullable: false),
                    IsSOH = table.Column<bool>(type: "bit", nullable: false),
                    IsRace = table.Column<bool>(type: "bit", nullable: false),
                    IsStock = table.Column<bool>(type: "bit", nullable: false),
                    IsOLR = table.Column<bool>(type: "bit", nullable: false),
                    IsBP = table.Column<bool>(type: "bit", nullable: false),
                    ImagePath = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    SireRingYr = table.Column<int>(type: "int", nullable: false),
                    SireRingMark = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    SireRingNo = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DamRingYr = table.Column<int>(type: "int", nullable: false),
                    DamRingMark = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    DamRingNo = table.Column<string>(type: "nvarchar(max)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Birds", x => x.BirdId);
                });

            migrationBuilder.CreateTable(
                name: "BloodLines",
                columns: table => new
                {
                    BloodLineId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    BloodLineName = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    Status = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BloodLines", x => x.BloodLineId);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Birds");

            migrationBuilder.DropTable(
                name: "BloodLines");
        }
    }
}
