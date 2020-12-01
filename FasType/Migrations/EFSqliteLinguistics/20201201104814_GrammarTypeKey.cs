using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations.EFSqliteLinguistics
{
    public partial class GrammarTypeKey : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "Name",
                table: "GrammarTypes",
                type: "TEXT",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddPrimaryKey(
                name: "PK_GrammarTypes",
                table: "GrammarTypes",
                column: "Name");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropPrimaryKey(
                name: "PK_GrammarTypes",
                table: "GrammarTypes");

            migrationBuilder.DropColumn(
                name: "Name",
                table: "GrammarTypes");
        }
    }
}
