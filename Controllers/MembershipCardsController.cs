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
    public class MembershipCardsController : Controller
    {
        private Model1 db = new Model1();

        [HttpPost]
        public ActionResult CreatePartial(MembershipCard mc)
        {
            var mcname = mc.Name;
            var mcjob = mc.Job;
            var mccreate = mc.CreatedDate;
            var mcexpire = mc.ExpiredDate;
            var mcadmin = mc.CreatedBy;
            DateTime expiredDate = mc.ExpiredDate ?? DateTime.Now;
            if (!string.IsNullOrEmpty(mcname))
            {
                var compare = DateTime.Compare(DateTime.Now,expiredDate);
                if (compare < 0)
                {
                    mc.Status = true;
                }
                else mc.Status = false;
                db.MembershipCards.Add(mc);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return PartialView("CreatePartial");
        }
        [HttpGet]
        public ActionResult GetByID(int ID)
        {
            MembershipCard mc = db.MembershipCards.Find(ID);
            return PartialView("EditByID", mc);
        }
        public ActionResult DeleteByID(int ID)
        {
            MembershipCard mc = db.MembershipCards.Find(ID);
            return PartialView("DeleteByID", mc);
        }
        public ActionResult EditByID(MembershipCard mc)
        {
            var mcname = mc.Name;
            var mcjob = mc.Job;
            var mccreate = mc.CreatedDate;
            var mcexpire = mc.ExpiredDate;
            var mcadmin = mc.CreatedBy;

            if (!string.IsNullOrEmpty(mcname))
            {
                db.Entry(mc).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return PartialView("EditByID");
        }
        [HttpPost]
        public ActionResult DeletePartial(int ID)
        {
            MembershipCard mc = db.MembershipCards.Find(ID);
            db.MembershipCards.Remove(mc);
            db.SaveChanges();
            return RedirectToAction("Index");
        }
        // GET: MembershipCards
        public ActionResult Index(string sortOrder, string search, string currentFilter, int? page) 
        {
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NameSortParm = String.IsNullOrEmpty(sortOrder) ? "name_desc" : "";
            ViewBag.JobSortParm = String.IsNullOrEmpty(sortOrder) ? "job_desc" : "";
            ViewBag.BorrowSortParm = String.IsNullOrEmpty(sortOrder) ? "borrow_desc" : "";
            ViewBag.PendingSortParm = String.IsNullOrEmpty(sortOrder) ? "pend_desc" : "";
            ViewBag.BorrowDateSortParm = String.IsNullOrEmpty(sortOrder) ? "brdate_desc" : "";
            ViewBag.ExpiredDateSortParm = String.IsNullOrEmpty(sortOrder) ? "exdate_desc" : "";
            ViewBag.IDSortParm = String.IsNullOrEmpty(sortOrder) ? "id_desc" : "";
            if (search != null)
            {
                page = 1;
            }
            else
            {
                search = currentFilter;
            }
            ViewBag.CurrentFilter = search;
            var cards = from mc in db.MembershipCards
                        select mc;
            if (!String.IsNullOrEmpty(search))
            {
                cards = cards.Where(s => s.Name.Contains(search)
                                       || s.Job.Contains(search));
            }
            switch (sortOrder)
            {
                case "name_desc":
                    cards = cards.OrderBy(s => s.Name);
                    break;
                case "job_desc":
                    cards = cards.OrderBy(s => s.Job);
                    break;
                case "borrow_desc":
                    cards = cards.OrderBy(s => s.CreatedDate);
                    break;
                case "pend_desc":
                    cards = cards.OrderBy(s => s.ExpiredDate);
                    break;
                case "brdate_desc":
                    cards = cards.OrderBy(s => s.PendingCount);
                    break;
                case "exdate_desc":
                    cards = cards.OrderBy(s => s.RentedCount);
                    break;
                case "id_desc":
                    cards = cards.OrderBy(s => s.ID);
                    break;
                default:
                    cards = cards.OrderBy(s => s.ID);
                    break;
            }
            int pageSize = 5;
            int pageNumber = (page ?? 1);
            return View(cards.ToPagedList(pageNumber,pageSize));
        }

        // GET: MembershipCards/Details/5
        public ActionResult Details(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MembershipCard membershipCard = db.MembershipCards.Find(id);
            if (membershipCard == null)
            {
                return HttpNotFound();
            }
            return View(membershipCard);
        }

        // GET: MembershipCards/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: MembershipCards/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name,Job,CreatedDate,ExpiredDate,CreatedBy,ModifiedBy,RentedCount,PendingCount,Status")] MembershipCard membershipCard)
        {
            if (ModelState.IsValid)
            {
                db.MembershipCards.Add(membershipCard);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(membershipCard);
        }

        // GET: MembershipCards/Edit/5
        public ActionResult Edit(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MembershipCard membershipCard = db.MembershipCards.Find(id);
            if (membershipCard == null)
            {
                return HttpNotFound();
            }
            return View(membershipCard);
        }

        // POST: MembershipCards/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name,Job,CreatedDate,ExpiredDate,CreatedBy,ModifiedBy,RentedCount,PendingCount,Status")] MembershipCard membershipCard)
        {
            if (ModelState.IsValid)
            {
                db.Entry(membershipCard).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(membershipCard);
        }

        // GET: MembershipCards/Delete/5
        public ActionResult Delete(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            MembershipCard membershipCard = db.MembershipCards.Find(id);
            if (membershipCard == null)
            {
                return HttpNotFound();
            }
            return View(membershipCard);
        }

        // POST: MembershipCards/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(long id)
        {
            MembershipCard membershipCard = db.MembershipCards.Find(id);
            db.MembershipCards.Remove(membershipCard);
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

            var list = db.MembershipCards.ToList();

            Sheet.Cells["A1"].Value = "Mã thành viên";
            Sheet.Cells["B1"].Value = "Tên thành viên";
            Sheet.Cells["C1"].Value = "Nghề nghiệp";
            Sheet.Cells["D1"].Value = "Số sách đang mượn";
            Sheet.Cells["E1"].Value = "Số sách đã mượn";
            Sheet.Cells["F1"].Value = "Ngày tạo thẻ";
            Sheet.Cells["G1"].Value = "Ngày hết hạn";
            Sheet.Cells["H1"].Value = "Người tạo";
            using (var range = Sheet.Cells["A1:L1"])
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
                foreach(var item in list)
                {
                    countPending += item.PendingCount;
                    countBorrow += item.RentedCount;
                    Sheet.Cells[string.Format("A{0}", row)].Value = item.ID;
                    Sheet.Cells[string.Format("B{0}", row)].Value = item.Name;
                    Sheet.Cells[string.Format("C{0}", row)].Value = item.Job;
                    Sheet.Cells[string.Format("D{0}", row)].Value = item.PendingCount;
                    Sheet.Cells[string.Format("E{0}", row)].Value = item.RentedCount;
                    Sheet.Cells[string.Format("F{0}", row)].Value = item.CreatedDate.ToString();
                    Sheet.Cells[string.Format("G{0}", row)].Value = item.ExpiredDate.ToString();
                    Sheet.Cells[string.Format("H{0}", row)].Value = item.CreatedBy;
                    Sheet.Cells["A{0}:H{0}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    row++;
                }

                Sheet.Cells["J1"].Value = "Thống kê";
                Sheet.Cells["K1"].Value = "Tổng số sách đang mượn";
                Sheet.Cells["L1"].Value = "Tổng số sách đã mượn";
                Sheet.Cells["K2"].Value = countPending;
                Sheet.Cells["L2"].Value = countBorrow;

                



                Sheet.Cells["A:AZ"].AutoFitColumns();
                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.BinaryWrite(pck.GetAsByteArray());
                Response.End();
            }
        }
    }
}
