namespace QLTV.EF
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;
    using System.Linq.Expressions;
    using System.Web.Mvc;

    [Table("Book")]
    public partial class Book
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        
        public Book()
        {
            Borrows = new HashSet<Borrow>();
            this.Status = true;
        }

        public long ID { get; set; }

        [StringLength(50)]
        public string Name { get; set; }

        [StringLength(50)]
        public string Code { get; set; }

        public long? CategoryID { get; set; }

        [StringLength(250)]
        public string Description { get; set; }

        [StringLength(50)]
        public string Language { get; set; }

        public int? YearOfPublishment { get; set; }

        [StringLength(250)]
        public string Image { get; set; }

        [Required]
        [StringLength(50)]
        public string Author { get; set; }

        [StringLength(50)]
        public string Translator { get; set; }

        [StringLength(50)]
        public string Publisher { get; set; }

        public int? Quantity { get; set; }

        [StringLength(50)]
        public string StorageLocation { get; set; }

        public int? RentedCount { get; set; }

        [StringLength(50)]
        public string Performance { get; set; }

        public bool? Status { get; set; }

        public int? DisplayOrder { get; set; }

        public DateTime? CreatedDate { get; set; }

        [StringLength(50)]
        public string CreatedBy { get; set; }

        public DateTime? ModifiedDate { get; set; }

        [StringLength(50)]
        public string ModifiedBy { get; set; }

        public long? RentedBy { get; set; }

        public virtual BookCategory BookCategory { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Borrow> Borrows { get; set; }
        
    }
}
