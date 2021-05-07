using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using QLTV.EF;
using PagedList;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace QLTV.Controllers
{
    public class BorrowsController : Controller
    {
        private Model1 db = new Model1();

        [HttpPost]
        public ActionResult CreatePartial(Borrow br)
        {
            var bradmin = br.Admin;
            if (!string.IsNullOrEmpty(bradmin))
            {
                db.Borrows.Add(br);
                Book b = db.Books.Find(br.BookID);
                MembershipCard mc = db.MembershipCards.Find(br.MemberID);
                b.RentedCount++;
                mc.RentedCount++;
                mc.PendingCount++;
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return PartialView("CreatePartial");
        }
        [HttpGet]
        public ActionResult GetByID(int ID)
        {
            Borrow b = db.Borrows.Find(ID);
            return PartialView("EditByID", b);
        }
        public ActionResult DeleteByID(int ID)
        {
            Borrow b = db.Borrows.Find(ID);
            return PartialView("DeleteByID", b);
        }
        public ActionResult EditByID(Borrow b)
        {
            var bdes = b.Admin;

            if (!string.IsNullOrEmpty(bdes))
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
            Borrow b = db.Borrows.Find(ID);
            Book book = db.Books.Find(b.BookID);
            MembershipCard mc = db.MembershipCards.Find(b.MemberID);
            book.RentedCount--;
            mc.PendingCount--;
            db.Borrows.Remove(b);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        // GET: Borrows
        public ActionResult Index(string sortOrder,string search, string currentFilter, int? page)
        {
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewBag.MemberSortParm = String.IsNullOrEmpty(sortOrder) ? "mem_desc" : "";
            ViewBag.DateSortParm = String.IsNullOrEmpty(sortOrder) ? "Date" : "";
            ViewBag.ExpiredDateSortParm = String.IsNullOrEmpty(sortOrder) ? "date_desc" : "";
            ViewBag.bIDSortParm = String.IsNullOrEmpty(sortOrder) ? "bookid_desc" : "";
            ViewBag.mIDSortParm = String.IsNullOrEmpty(sortOrder) ? "memid_desc" : "";
            if (search != null)
            {
                page = 1;
            }
            else
            {
                search = currentFilter;
            }
            ViewBag.CurrentFilter = search;
            var borrows = from s in db.Borrows
                           select s;
            if (!String.IsNullOrEmpty(search))
            {
                borrows = borrows.Where(s => s.Book.Name.Contains(search)
                                       || s.MembershipCard.Name.Contains(search) );
            }
            switch (sortOrder)
            {
                case "name_desc":
                    borrows = borrows.OrderBy(s => s.Book.Name);
                    break;
                case "mem_desc":
                    borrows = borrows.OrderBy(s => s.MembershipCard.Name);
                    break;
                case "Date":
                    borrows = borrows.OrderBy(s => s.BorrowDate);
                    break;
                case "date_desc":
                    borrows = borrows.OrderBy(s => s.ExpiredDate);
                    break;
                case "memid_desc":
                    borrows = borrows.OrderBy(s => s.MemberID);
                    break;
                case "bookid_desc":
                    borrows = borrows.OrderBy(s => s.BookID);
                    break;
                default:
                    borrows = borrows.OrderBy(s => s.ID);
                    break;
            }
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            
            
            return View(borrows.ToPagedList(pageNumber,pageSize));
        }

        // GET: Borrows/Details/5
        public ActionResult Details(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Borrow borrow = db.Borrows.Find(id);
            if (borrow == null)
            {
                return HttpNotFound();
            }
            return View(borrow);
        }

        // GET: Borrows/Create
        public ActionResult Create()
        {
            ViewBag.BookID = new SelectList(db.Books, "ID", "Name");
            ViewBag.MemberID = new SelectList(db.MembershipCards, "ID", "Name");
            return View();
        }

        // POST: Borrows/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,BookID,MemberID,BookName,Author,Description,Note,BorrowDate,ExpiredDate,Admin")] Borrow borrow)
        {
            if (ModelState.IsValid)
            {
                db.Borrows.Add(borrow);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.BookID = new SelectList(db.Books, "ID", "Name", borrow.BookID);
            ViewBag.MemberID = new SelectList(db.MembershipCards, "ID", "Name", borrow.MemberID);
            return View(borrow);
        }

        // GET: Borrows/Edit/5
        public ActionResult Edit(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Borrow borrow = db.Borrows.Find(id);
            if (borrow == null)
            {
                return HttpNotFound();
            }
            ViewBag.BookID = new SelectList(db.Books, "ID", "Name", borrow.BookID);
            ViewBag.MemberID = new SelectList(db.MembershipCards, "ID", "Name", borrow.MemberID);
            return View(borrow);
        }

        // POST: Borrows/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,BookID,MemberID,BookName,Author,Description,Note,BorrowDate,ExpiredDate,Admin")] Borrow borrow)
        {
            if (ModelState.IsValid)
            {
                db.Entry(borrow).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.BookID = new SelectList(db.Books, "ID", "Name", borrow.BookID);
            ViewBag.MemberID = new SelectList(db.MembershipCards, "ID", "Name", borrow.MemberID);
            return View(borrow);
        }

        // GET: Borrows/Delete/5
        public ActionResult Delete(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Borrow borrow = db.Borrows.Find(id);
            if (borrow == null)
            {
                return HttpNotFound();
            }
            return View(borrow);
        }

        // POST: Borrows/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(long id)
        {
            Borrow borrow = db.Borrows.Find(id);
            db.Borrows.Remove(borrow);
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

            ExcelPackage pck = new ExcelPackage(new System.IO.FileInfo(fileName));

            var Sheet = pck.Workbook.Worksheets.Add("Report");

            var list = db.Borrows.ToList();

            Sheet.Cells["A1"].Value = "Mã đơn";
            Sheet.Cells["B1"].Value = "ID sách";
            Sheet.Cells["C1"].Value = "ID Thành viên";
            Sheet.Cells["D1"].Value = "Tên sách";
            Sheet.Cells["E1"].Value = "Tên thành viên";
            Sheet.Cells["F1"].Value = "Ngày mượn";
            Sheet.Cells["G1"].Value = "Ngày hết hạn";
            Sheet.Cells["H1"].Value = "Ghi chú";
            
            
            using (var range = Sheet.Cells["A1:H1"])
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
            
            if (true)
            {
                foreach (var item in list)
                {
                    
                    Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                    Sheet.Cells[string.Format("B{0}", row)].Value = item.BookID;
                    Sheet.Cells[string.Format("C{0}", row)].Value = item.MemberID;
                    Sheet.Cells[string.Format("D{0}", row)].Value = item.Book.Name;
                    Sheet.Cells[string.Format("E{0}", row)].Value = item.MembershipCard.Name;
                    Sheet.Cells[string.Format("F{0}", row)].Value = item.BorrowDate.ToString();
                    Sheet.Cells[string.Format("G{0}", row)].Value = item.ExpiredDate.ToString();
                    Sheet.Cells[string.Format("H{0}", row)].Value = item.Description;
                    
                    Sheet.Cells["A{0}:H{0}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    row++;
                }

                





                Sheet.Cells["A:AZ"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=BorrowsReport.xlsx");
                Response.BinaryWrite(pck.GetAsByteArray());
                Response.End();
            }
        }
    }
}
