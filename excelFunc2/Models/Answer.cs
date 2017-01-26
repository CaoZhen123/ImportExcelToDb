using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace excelFunc2.Models
{
    public class Answer
    {
        public int ID { get; set; }

        public string Content { get; set; }

        public string Explaination { get; set; }

        public int Answer_Flag { get; set; }

        public Question Question { get; set; }
        public int QuestionId { get; set; }

    }
}