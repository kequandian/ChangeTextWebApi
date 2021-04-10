using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChangeText.WebApi.Models
{
    public class WordToPdfData
    {
        public string Sourcedocx { get; set; }
        public string Targetpdf { get; set; }
    }
}