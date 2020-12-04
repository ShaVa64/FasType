using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class RemoveOtherShortForms : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ShortGenderForm",
                table: "Abbreviations");

            migrationBuilder.DropColumn(
                name: "ShortGenderPluralForm",
                table: "Abbreviations");

            migrationBuilder.DropColumn(
                name: "ShortPluralForm",
                table: "Abbreviations");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "ShortGenderForm",
                table: "Abbreviations",
                type: "TEXT",
                maxLength: 50,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ShortGenderPluralForm",
                table: "Abbreviations",
                type: "TEXT",
                maxLength: 50,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ShortPluralForm",
                table: "Abbreviations",
                type: "TEXT",
                maxLength: 50,
                nullable: true);
        }
    }
}
