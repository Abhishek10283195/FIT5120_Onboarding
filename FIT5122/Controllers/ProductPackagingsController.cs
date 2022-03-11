using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Data.OleDb;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using FIT5122.Models;
using LinqToExcel;

namespace FIT5122.Controllers
{
    public class ProductPackagingsController : Controller
    {
        private ProductPackagingContainer db = new ProductPackagingContainer();

        // GET: ProductPackagings
        public ActionResult Index()
        {
            return View(db.ProductPackagings.ToList());
        }
        public FileResult DownloadExcel()
        {
            string path = "~Uploads/openfoodfacts_export.xlsx";
            return File(path, "application/vnd.ms-excel", "openfoodfacts_export.xlsx");
        }

        [HttpPost]
        public JsonResult UploadExcel(ProductPackagings users, HttpPostedFileBase FileUpload)
        {

            List<string> data = new List<string>();
            if (FileUpload != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");
                if (FileUpload.ContentType == "application/vnd.ms-excel" || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = FileUpload.FileName;
                    string targetpath = Server.MapPath("~/Doc/");
                    FileUpload.SaveAs(targetpath + filename);
                    string pathToExcelFile = targetpath + filename;
                    var connectionString = "";
                    if (filename.EndsWith(".xls"))
                    {
                        connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", pathToExcelFile);
                    }
                    else if (filename.EndsWith(".xlsx"))
                    {
                        connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", pathToExcelFile);
                    }

                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "ExcelTable");
                    DataTable dtable = ds.Tables["ExcelTable"];
                    string sheetName = "Sheet1";
                    var excelFile = new ExcelQueryFactory(pathToExcelFile);
                    var artistAlbums = from a in excelFile.Worksheet<ProductPackagings>(sheetName) select a;
                    foreach (var a in artistAlbums)
                    {
                        try
                        {
                            if (a.Packaging != "" && a.Productname != "")
                            {
                                ProductPackagings TU = new ProductPackagings();
                                TU.Productname = a.Productname;
                                TU.Packaging = a.Packaging;
                                db.ProductPackagings.Add(TU);
                                db.SaveChanges();
                            }
                            else
                            {
                                data.Add("<ul>");
                                if (a.Productname == "" || a.Productname == null) data.Add("<li> name is required</li>");
                                if (a.Packaging == "" || a.Packaging == null) data.Add("<li> Address is required</li>");
                                data.Add("</ul>");
                                data.ToArray();
                                return Json(data, JsonRequestBehavior.AllowGet);
                            }
                        }
                        catch (DbEntityValidationException ex)
                        {
                            foreach (var entityValidationErrors in ex.EntityValidationErrors)
                            {
                                foreach (var validationError in entityValidationErrors.ValidationErrors)
                                {
                                    Response.Write("Property: " + validationError.PropertyName + " Error: " + validationError.ErrorMessage);
                                }
                            }
                        }
                    }
                    //deleting excel file from folder
                    if ((System.IO.File.Exists(pathToExcelFile)))
                    {
                        System.IO.File.Delete(pathToExcelFile);
                    }
                    return Json("success", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //alert message for invalid file format
                    data.Add("<ul>");
                    data.Add("<li>Only Excel file format is allowed</li>");
                    data.Add("</ul>");
                    data.ToArray();
                    return Json(data, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                data.Add("<ul>");
                if (FileUpload == null) data.Add("<li>Please choose Excel file</li>");
                data.Add("</ul>");
                data.ToArray();
                return Json(data, JsonRequestBehavior.AllowGet);
            }
        }
        // GET: ProductPackagings/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ProductPackagings productPackagings = db.ProductPackagings.Find(id);
            if (productPackagings == null)
            {
                return HttpNotFound();
            }
            return View(productPackagings);
        }

        // GET: ProductPackagings/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ProductPackagings/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "Id,Productname,Packaging")] ProductPackagings productPackagings)
        {
            if (ModelState.IsValid)
            {
                db.ProductPackagings.Add(productPackagings);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(productPackagings);
        }

        // GET: ProductPackagings/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ProductPackagings productPackagings = db.ProductPackagings.Find(id);
            if (productPackagings == null)
            {
                return HttpNotFound();
            }
            return View(productPackagings);
        }

        // POST: ProductPackagings/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,Productname,Packaging")] ProductPackagings productPackagings)
        {
            if (ModelState.IsValid)
            {
                db.Entry(productPackagings).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(productPackagings);
        }

        // GET: ProductPackagings/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ProductPackagings productPackagings = db.ProductPackagings.Find(id);
            if (productPackagings == null)
            {
                return HttpNotFound();
            }
            return View(productPackagings);
        }

        // POST: ProductPackagings/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            ProductPackagings productPackagings = db.ProductPackagings.Find(id);
            db.ProductPackagings.Remove(productPackagings);
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
