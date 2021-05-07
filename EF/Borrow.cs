namespace QLTV.EF
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Borrow")]
    public partial class Borrow
    {
        public Borrow()
        {
         
        }
        public long ID { get; set; }

        public long? BookID { get; set; }

        public long? MemberID { get; set; }

        [StringLength(50)]
        public string BookName { get; set; }

        [StringLength(50)]
        public string Author { get; set; }

        public string Description { get; set; }

        [StringLength(50)]
        public string Note { get; set; }

        public DateTime? BorrowDate { get; set; }

        public DateTime? ExpiredDate { get; set; }

        [StringLength(50)]
        public string Admin { get; set; }

        public virtual Book Book { get; set; }

        public virtual MembershipCard MembershipCard { get; set; }
    }
}
