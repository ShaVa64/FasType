using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace FasType.Migrations.EFSqliteLinguistics
{
    public partial class AddAbbreviationMethodKey : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<Guid>(
                name: "Key",
                table: "AbbreviationMethods",
                type: "TEXT",
                nullable: false,
                defaultValue: new Guid("00000000-0000-0000-0000-000000000000"));

            migrationBuilder.AddPrimaryKey(
                name: "PK_AbbreviationMethods",
                table: "AbbreviationMethods",
                column: "Key");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropPrimaryKey(
                name: "PK_AbbreviationMethods",
                table: "AbbreviationMethods");

            migrationBuilder.DropColumn(
                name: "Key",
                table: "AbbreviationMethods");
        }
    }
}
