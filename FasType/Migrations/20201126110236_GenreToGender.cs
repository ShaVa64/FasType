using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class GenreToGender : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "GenreForm",
                table: "Abbreviations",
                newName: "GenderForm");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "GenderForm",
                table: "Abbreviations",
                newName: "GenreForm");
        }
    }
}
