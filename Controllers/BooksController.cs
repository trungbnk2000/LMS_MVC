using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PagedList;
using QLTV.EF;

namespace QLTV.Controllers
{
    public class BooksController : Controller
    {
        private Model1 db = new Model1();

        [HttpPost]
        public ActionResult CreatePartial(Book b)
        {
            var bname = b.Name;

            if (!string.IsNullOrEmpty(bname))
            {
                db.Books.Add(b);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return PartialView("CreatePartial");
        }
        
        [HttpGet]
        public ActionResult GetByID(int ID)
        {
            Book b = db.Books.Find(ID);
            return PartialView("EditByID", b);
        }
        public ActionResult DeleteByID(int ID)
        {
            Book b = db.Books.Find(ID);
            return PartialView("DeleteByID", b);
        }
        public ActionResult EditByID(Book b)
        {
            var bname = b.Name;

            if (!string.IsNullOrEmpty(bname))
            {
                db.Entry(b).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return PartialView("EditByID");
        }
        [HttpPost]
        public ActionResult DeletePartial(int ID)
        {
            Book b = db.Books.Find(ID);
            db.Books.Remove(b);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        // GET: Books

        public ActionResult Index(string sortOrder, string search, string currentFilter, int? page)
        {
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NameSort = String.IsNullOrEmpty(sortOrder) ? "nameSort" : "";
            ViewBag.AuthorSort = String.IsNullOrEmpty(sortOrder) ? "authorSort" : "";
            ViewBag.TypeSort = String.IsNullOrEmpty(sortOrder) ? "typeSort" : "";
            ViewBag.RentSort = String.IsNullOrEmpty(sortOrder) ? "rentSort" : "";
            ViewBag.QuantitySort = String.IsNullOrEmpty(sortOrder) ? "quantitySort" : "";
            if (search != null)
            {
                page = 1;
            }
            else
            {
                search = currentFilter;
            }
            ViewBag.CurrentFilter = search;
            var books = from b in db.Books
                        select b;
            if (!String.IsNullOrEmpty(search))
            {
                books = books.Where(b => b.Name.Contains(search) || b.Author.Contains(search));
            }
            switch (sortOrder)
            {
                case "nameSort":
                    books = books.OrderBy(s => s.Name);
                    break;
                case "authorSort":
                    books = books.OrderBy(s => s.Author);
                    break;
                case "typeSort":
                    books = books.OrderBy(s => s.BookCategory);
                    break;
                case "rentSort":
                    books = books.OrderBy(s => s.RentedCount);
                    break;
                case "quantitySort":
                    books = books.OrderBy(s => s.Quantity);
                    break;
                default:
                    books = books.OrderBy(s => s.ID);
                    break;
            }
            int pageSize = 5;
            int pageNumber = (page ?? 1);

            return View(books.ToPagedList(pageNumber,pageSize));
        }
        
        // GET: Books/Details/5
        public ActionResult Details(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Book book = db.Books.Find(id);
            if (book == null)
            {
                return HttpNotFound();
            }
            return View(book);
        }
        
        // GET: Books/Create
        public ActionResult Create()
        {
            ViewBag.CategoryID = new SelectList(db.BookCategories, "ID", "Name");
            return View();
        }
        
        // POST: Books/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name,Code,CategoryID,Description,Language,YearOfPublishment,Image,Author,Translator,Publisher,Quantity,StorageLocation,RentedCount,Performance,Status,DisplayOrder,CreatedDate,CreatedBy,ModifiedDate,ModifiedBy,RentedBy")] Book book)
        {
            if (ModelState.IsValid)
            {
                db.Books.Add(book);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.CategoryID = new SelectList(db.BookCategories, "ID", "Name", book.CategoryID);
            return View(book);
        }

        // GET: Books/Edit/5
        public ActionResult Edit(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Book book = db.Books.Find(id);
            if (book == null)
            {
                return HttpNotFound();
            }
            ViewBag.CategoryID = new SelectList(db.BookCategories, "ID", "Name", book.CategoryID);
            return View(book);
        }

        // POST: Books/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name,Code,CategoryID,Description,Language,YearOfPublishment,Image,Author,Translator,Publisher,Quantity,StorageLocation,RentedCount,Performance,Status,DisplayOrder,CreatedDate,CreatedBy,ModifiedDate,ModifiedBy,RentedBy")] Book book)
        {
            if (ModelState.IsValid)
            {
                db.Entry(book).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.CategoryID = new SelectList(db.BookCategories, "ID", "Name", book.CategoryID);
            return View(book);
        }

        // GET: Books/Delete/5
        public ActionResult Delete(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Book book = db.Books.Find(id);
            if (book == null)
            {
                return HttpNotFound();
            }
            return View(book);
        }
        
        // POST: Books/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(long id)
        {
            Book book = db.Books.Find(id);
            db.Books.Remove(book);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
        [HttpGet]
        public void ExportExcel()
        {
            String fileName = Server.MapPath("/") + "\\export\\myExport.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(fileName));

            var Sheet = pck.Workbook.Worksheets.Add("Report");

            var list = db.Books.ToList();

            Sheet.Cells["A1"].Value = "Mã sách";
            Sheet.Cells["B1"].Value = "Tên sách";
            Sheet.Cells["C1"].Value = "Tác giả";
            Sheet.Cells["D1"].Value = "Số sách đang mượn";
            Sheet.Cells["E1"].Value = "Tổng số sách";
            Sheet.Cells["F1"].Value = "Thể loại sách";
            Sheet.Cells["G1"].Value = "Người dịch";
            Sheet.Cells["H1"].Value = "Nhà xuất bản";
            Sheet.Cells["I1"].Value = "Vị trí lưu giữ";
            Sheet.Cells["J1"].Value = "Tình trạng";
            using (var range = Sheet.Cells["A1:O1"])
            {
                // Set PatternType
                range.Style.Fill.PatternType = ExcelFillStyle.DarkGray;
                // Set Màu cho Background
                range.Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                // Canh giữa cho các text
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                // Set Font cho text  trong Range hiện tại
                range.Style.Font.SetFromFont(new Font("Calibri", 12));
                // Set Border
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                // Set màu ch Border
                range.Style.Border.Bottom.Color.SetColor(Color.Blue);

                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            int row = 2;
            int? countPending = 0, countBorrow = 0;
            if (true)
            {
                foreach (var item in list)
                {
                    countPending += item.RentedCount;
                    countBorrow += item.Quantity;
                    Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                    Sheet.Cells[string.Format("B{0}", row)].Value = item.Name;
                    Sheet.Cells[string.Format("C{0}", row)].Value = item.Author;
                    Sheet.Cells[string.Format("D{0}", row)].Value = item.RentedCount;
                    Sheet.Cells[string.Format("E{0}", row)].Value = item.Quantity;
                    Sheet.Cells[string.Format("F{0}", row)].Value = item.CategoryID;
                    Sheet.Cells[string.Format("G{0}", row)].Value = item.Translator;
                    Sheet.Cells[string.Format("H{0}", row)].Value = item.Publisher;
                    Sheet.Cells[string.Format("I{0}", row)].Value = item.StorageLocation;
                    Sheet.Cells[string.Format("J{0}", row)].Value = item.Status;
                    Sheet.Cells["A{0}:J{0}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    row++;
                }

                Sheet.Cells["M1"].Value = "Thống kê";
                Sheet.Cells["N1"].Value = "Tổng số sách đang mượn";
                Sheet.Cells["O1"].Value = "Tổng số sách";
                Sheet.Cells["N2"].Value = countPending;
                Sheet.Cells["O2"].Value = countBorrow;





                Sheet.Cells["A:AZ"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=BooksReport.xlsx");
                Response.BinaryWrite(pck.GetAsByteArray());
                Response.End();
            }
        }
    }
}
