using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations.EFSqliteLinguistics
{
    public partial class LinguisticsInit : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "AbbreviationMethods",
                columns: table => new
                {
                    ShortForm = table.Column<string>(type: "TEXT", nullable: true),
                    FullForm = table.Column<string>(type: "TEXT", nullable: true),
                    Position = table.Column<int>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                });

            migrationBuilder.CreateTable(
                name: "GrammarTypes",
                columns: table => new
                {
                    Repr = table.Column<string>(type: "TEXT", nullable: true),
                    Position = table.Column<int>(type: "INTEGER", nullable: false)
                },
                constraints: table =>
                {
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AbbreviationMethods");

            migrationBuilder.DropTable(
                name: "GrammarTypes");
        }
    }
}
