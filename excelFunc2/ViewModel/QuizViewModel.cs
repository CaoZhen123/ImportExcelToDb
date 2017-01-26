using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using excelFunc2.Models;

namespace excelFunc2.ViewModel
{
    public class QuizViewModel
    {
        public Question question { get; set; }
        public List<Answer> answers { get; set; }
    }
}