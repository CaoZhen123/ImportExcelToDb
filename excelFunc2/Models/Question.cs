using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace excelFunc2.Models
{
    public class Question
    {
        int ID { get; set; }

        string Content { get; set; }

        int Difficulty { get; set; }

        int NumberOfWrongGlobal { get; set; }

        int NumberOfCorrectGlobal { get; set; }

    }
}