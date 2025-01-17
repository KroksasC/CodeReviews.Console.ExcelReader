using Microsoft.EntityFrameworkCore;
using System.Configuration;

namespace ExcelReader.KroksasC
{
    internal class ExcelContext : DbContext
    {
        public DbSet<Person> Persons { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
            => optionsBuilder.UseSqlServer(ConfigurationManager.ConnectionStrings["ExcelReader"].ConnectionString);
    }
}
