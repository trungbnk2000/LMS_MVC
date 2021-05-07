using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;

namespace QLTV.EF
{
    public partial class Model1 : DbContext
    {
        public Model1()
            : base("name=DBContext")
        {
        }

        public virtual DbSet<Book> Books { get; set; }
        public virtual DbSet<BookCategory> BookCategories { get; set; }
        public virtual DbSet<Borrow> Borrows { get; set; }
        public virtual DbSet<MembershipCard> MembershipCards { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<BookCategory>()
                .HasMany(e => e.Books)
                .WithOptional(e => e.BookCategory)
                .HasForeignKey(e => e.CategoryID);

            modelBuilder.Entity<MembershipCard>()
                .HasMany(e => e.Borrows)
                .WithOptional(e => e.MembershipCard)
                .HasForeignKey(e => e.MemberID);
        }
    }
}
