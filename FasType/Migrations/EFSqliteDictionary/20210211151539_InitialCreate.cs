using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations.EFSqliteDictionary
{
    public partial class InitialCreate : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Dictionary",
                columns: table => new
                {
                    FullForm = table.Column<string>(type: "TEXT", nullable: false),
                    Others = table.Column<string>(type: "TEXT", nullable: true),
                    AllForms = table.Column<string>(type: "TEXT", nullable: true),
                    Discriminator = table.Column<string>(type: "TEXT", nullable: false),
                    GenderForm = table.Column<string>(type: "TEXT", nullable: true),
                    PluralForm = table.Column<string>(type: "TEXT", nullable: true),
                    GenderPluralForm = table.Column<string>(type: "TEXT", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Dictionary", x => x.FullForm);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Dictionary");
        }
    }
}
