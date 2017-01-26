using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace excelFunc2.Models
{
    public class Question
    {
        public int ID { get; set; }

        public string Content { get; set; }

        public int Difficulty { get; set; }

        public int NumberOfWrongGlobal { get; set; }

        public int NumberOfCorrectGlobal { get; set; }

    }
}