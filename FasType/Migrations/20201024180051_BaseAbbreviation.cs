using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class BaseAbbreviation : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Abbreviations",
                columns: table => new
                {
                    Key = table.Column<Guid>(nullable: false),
                    ShortForm = table.Column<string>(type: "varchar(50)", nullable: false),
                    FullForm = table.Column<string>(type: "varchar(50)", nullable: false),
                    Discriminator = table.Column<string>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Abbreviations", x => x.Key);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Abbreviations");
        }
    }
}
