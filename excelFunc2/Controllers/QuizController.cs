using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace excelFunc2.Controllers
{
    public class QuizController : Controller
    {
        // GET: Quiz
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Import(HttpPostedFileBase excelfile) {
            if (excelfile == null || excelfile.ContentLength == 0)                              // did not attach file
            {
                ViewBag.Error = "Please select an excel file";
                return View("Index");
            }   
            else {                                                                               // attached file
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))   // attached a correct excel file
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
                    
                    return View("Success");
                }
                else {                                                                           // attached uncorrect excel file
                    ViewBag.Error = "File type is incorrect";
                    return View("Index");
                }
            }
        }
    }
}