using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class AddGenderPlural : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "GenderPluralForm",
                table: "Abbreviations",
                type: "varchar(50)",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "GenderPluralForm",
                table: "Abbreviations");
        }
    }
}
