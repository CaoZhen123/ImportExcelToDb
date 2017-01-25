using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using excelFunc2.Models;

namespace excelFunc2.Controllers
{
    public class CustomerController : Controller
    {
        private ApplicationDbContext _context;
        public CustomerController()
        {
            _context = new ApplicationDbContext();
        }

        protected override void Dispose(bool disposing)
        {
            //base.Dispose(disposing);
            _context.Dispose();
        }
        // GET: Customer
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile) {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select a excel file";
                return View("Index");
            }
            else {

                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);

                    //Read data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<Customer> listCustomers = new List<Customer>();
                    for (int row = 1; row <= range.Rows.Count; row++) {
                        Customer cus = new Customer();
                        cus.Id = ((Excel.Range)range.Cells[row, 1]).Text;
                        cus.Name = ((Excel.Range)range.Cells[row, 2]).Text;
                        cus.Age = ((Excel.Range)range.Cells[row, 3]).Text;
                        _context.Customers.Add(cus);
                        listCustomers.Add(cus);

                    }
                    _context.SaveChanges();
                    ViewBag.ListCustomers = listCustomers;
                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect";
                    return View("Index");
                }
            }
        }
            
    }
}