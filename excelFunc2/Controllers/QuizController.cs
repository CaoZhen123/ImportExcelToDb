using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using excelFunc2.ViewModel;
using excelFunc2.Models;

namespace excelFunc2.Controllers
{
    public class QuizController : Controller
    {
        private ApplicationDbContext _context;

        public QuizController()
        {
            _context = new ApplicationDbContext();
        }

        protected override void Dispose(bool disposing)
        {
            _context.Dispose();
        }

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
                    for (int row = 1; row <= range.Rows.Count; row++) {
                        
                        Question que = new Question();   // extract question info and create the question
                        que.ID = Convert.ToInt32(((Excel.Range)range.Cells[row, 1]).Text);
                        que.Content = ((Excel.Range)range.Cells[row, 2]).Text;
                        que.Difficulty = Convert.ToInt32(((Excel.Range)range.Cells[row, 3]).Text);
                        que.NumberOfCorrectGlobal = 0;
                        que.NumberOfWrongGlobal = 0;
                        _context.Questions.Add(que);
                        for (int ans_num = 1; ans_num < 5; ans_num++) {
                            Answer ans = new Answer();
                            ans.Content = ((Excel.Range)range.Cells[row, ans_num+3]).Text;
                            if (ans_num == 1)
                            {
                                ans.Answer_Flag = 0;
                            }
                            else {
                                ans.Answer_Flag = 1;
                            }
                            ans.Explaination = "";
                            ans.QuestionId = Convert.ToInt32(((Excel.Range)range.Cells[row, 1]).Text);
                            _context.Answers.Add(ans);
                        }
                    }
                    _context.SaveChanges();
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