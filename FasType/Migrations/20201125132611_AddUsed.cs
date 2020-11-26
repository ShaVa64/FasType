using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations
{
    public partial class AddUsed : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<ulong>(
                name: "Used",
                table: "Abbreviations",
                type: "INTEGER",
                nullable: false,
                defaultValue: 0ul);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Used",
                table: "Abbreviations");
        }
    }
}
