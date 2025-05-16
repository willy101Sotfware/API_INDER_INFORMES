using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Models;

namespace API_INDER_INFORMES.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<FormularioInder> FormularioInder { get; set; }
        public DbSet<DisponibilidadHorariosInder> DisponibilidadHorariosInder { get; set; }



        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<DisponibilidadHorariosInder>()
                .HasKey(d => d.IdDis);

            modelBuilder.Entity<FormularioInder>()
                .HasKey(f => f.Id);
        }
    }
} 