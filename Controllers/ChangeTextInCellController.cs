using ChangeText.WebApi.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;

namespace ChangeText.WebApi.Controllers
{

    public class ChangeTextInCellController : ApiController
    {
		
        private ApiTools tool = new ApiTools();
		[HttpPost]
		public HttpResponseMessage ChangeText(ChangeTextData ctData)
		{

            try
            {
                WriteWordTable writeWordTable = new WriteWordTable();
                writeWordTable.ChangeTextInCell(ctData);
            }
            catch (Exception ex)
            {
                return tool.MsgFormat(ResponseCode.操作失败, "修改失败", ex.ToString());
            }

            return tool.MsgFormat(ResponseCode.成功, "修改成功", "0");
        }
	}

}