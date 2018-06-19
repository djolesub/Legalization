namespace Resenja_EF_Konzistentno_DIV
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class Resenja : DbContext
    {
        public Resenja()
            : base("connection parameters")
        {
        }

        public virtual DbSet<RR> RR { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }
    }
}
