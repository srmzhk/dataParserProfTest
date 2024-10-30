using DataParserProfTest.Model;
using System.Data.Entity;

namespace DataParserProfTest.Repositories
{
    public partial class AppDbContext : DbContext
    {
        public AppDbContext() : base("Server=(local);Database=ProfTestDP;Trusted_Connection=True;TrustServerCertificate=True;")
        {
        }

        public DbSet<Question> Questions { get; set; }
        public DbSet<Test> Tests { get; set; }
        public DbSet<Answer> Answers { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
        }
    }
}