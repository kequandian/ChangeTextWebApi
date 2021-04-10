using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChangeText.WebApi.Models
{
    public class ChangeTextData
    {
        public string DocUrl { get; set; }
        public string RowIndex { get; set; }
        public string ColumnIndex { get; set; }
        public string Content { get; set; }
    }
}