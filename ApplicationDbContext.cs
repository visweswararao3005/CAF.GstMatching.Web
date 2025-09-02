
using CAF.GstMatching.Web.Models;
using Microsoft.EntityFrameworkCore;

namespace CAF.GstMatching.Web
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options): base(options)
        {
        }

        public DbSet<ChatMessage> Tj_ChatMessages { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ChatMessage>()
            .HasKey(c => new { c.FromUserID, c.ToUserID, c.TaskID, c.CreatedOn });
        }
    }
}



