using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using QLTV.EF;

namespace QLTV.Controllers
{
    public class BookCategoriesController : Controller
    {
        private Model1 db = new Model1();

        [HttpPost]
        public ActionResult CreatePartial(BookCategory bc)
        {
            var bcname = bc.Name;
            if (!string.IsNullOrEmpty(bcname))
            {
                db.BookCategories.Add(bc);
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return PartialView("CreatePartialView");
        }

        [HttpGet]
        public ActionResult GetByID(int ID)
        {
            BookCategory bc = db.BookCategories.Find(ID);
            return PartialView("EditByID", bc);
        }
        public ActionResult DeleteByID(int ID)
        {
            BookCategory bc = db.BookCategories.Find(ID);
            return PartialView("DeleteByID", bc);
        }
        public ActionResult EditByID(BookCategory bc)
        {
            var bcname = bc.Name;

            if (!string.IsNullOrEmpty(bcname))
            {
                db.Entry(bc).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return PartialView("EditByID");
        }
        [HttpPost]
        public ActionResult DeletePartial(int ID)
        {
            BookCategory bc = db.BookCategories.Find(ID);
            db.BookCategories.Remove(bc);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        // GET: BookCategories
        public ActionResult Index()
        {
            return View(db.BookCategories.ToList());
        }

        // GET: BookCategories/Details/5
        public ActionResult Details(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BookCategory bookCategory = db.BookCategories.Find(id);
            if (bookCategory == null)
            {
                return HttpNotFound();
            }
            return View(bookCategory);
        }

        // GET: BookCategories/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: BookCategories/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name")] BookCategory bookCategory)
        {
            if (ModelState.IsValid)
            {
                db.BookCategories.Add(bookCategory);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(bookCategory);
        }

        // GET: BookCategories/Edit/5
        public ActionResult Edit(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BookCategory bookCategory = db.BookCategories.Find(id);
            if (bookCategory == null)
            {
                return HttpNotFound();
            }
            return View(bookCategory);
        }

        // POST: BookCategories/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name")] BookCategory bookCategory)
        {
            if (ModelState.IsValid)
            {
                db.Entry(bookCategory).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(bookCategory);
        }

        // GET: BookCategories/Delete/5
        public ActionResult Delete(long? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            BookCategory bookCategory = db.BookCategories.Find(id);
            if (bookCategory == null)
            {
                return HttpNotFound();
            }
            return View(bookCategory);
        }

        // POST: BookCategories/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(long id)
        {
            BookCategory bookCategory = db.BookCategories.Find(id);
            db.BookCategories.Remove(bookCategory);
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
    }
}
