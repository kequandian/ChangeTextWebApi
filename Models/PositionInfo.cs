using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docx.openxml.models
{
    public class PositionInfo
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 内容
        /// </summary>
        public string Content { get; set; }
    }

    public class FilePathInfo
    {
        public string FilePath { get; set; }
    }
}
