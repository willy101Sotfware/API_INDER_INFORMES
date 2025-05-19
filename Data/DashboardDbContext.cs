using Microsoft.EntityFrameworkCore;
using API_INDER_INFORMES.Models;

namespace API_INDER_INFORMES.Data
{
    public class DashboardDbContext : DbContext
    {
        public DashboardDbContext(DbContextOptions<DashboardDbContext> options)
            : base(options)
        {
        }

        public DbSet<Transaction> Transactions { get; set; }
        public DbSet<TransactionDetail> TransactionDetails { get; set; }
        public DbSet<StateTransaction> StateTransactions { get; set; }
        public DbSet<PayPad> PayPads { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

          
            modelBuilder.Entity<Transaction>().ToTable("Transactions", "business");
            modelBuilder.Entity<TransactionDetail>().ToTable("TransactionsDetail", "business");
            modelBuilder.Entity<StateTransaction>().ToTable("StateTransaction", "masters");
            modelBuilder.Entity<PayPad>().ToTable("PayPad", "business");
        }
    }
}
