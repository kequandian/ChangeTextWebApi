using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChangeText.WebApi.Models
{
    public class HtmlToPdfData
    {
        public string SourceHtml { get; set; }
        public string TargetPdf { get; set; }
    }
}