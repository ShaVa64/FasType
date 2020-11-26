using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class AddedPluralAndGenre : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "SimpleAbbreviations");

            migrationBuilder.AddColumn<string>(
                name: "GenreForm",
                table: "Abbreviations",
                type: "varchar(50)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "PluralForm",
                table: "Abbreviations",
                type: "varchar(50)",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "GenreForm",
                table: "Abbreviations");

            migrationBuilder.DropColumn(
                name: "PluralForm",
                table: "Abbreviations");

            migrationBuilder.CreateTable(
                name: "SimpleAbbreviations",
                columns: table => new
                {
                    Key = table.Column<Guid>(type: "TEXT", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SimpleAbbreviations", x => x.Key);
                });
        }
    }
}
