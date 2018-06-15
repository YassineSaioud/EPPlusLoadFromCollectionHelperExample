using Helpers;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace Controllers
{
    public class ExportController : Controller
    {
        public FileContentResult ExcelExportDownload()
        {
            var fileInfo = ExportHelper.BuildWorkbookFromCollection("Title",
                                                                     new Dictionary<string, string>
                                                                     {
                                                                         { "Header1", "PropertyName1" },
                                                                         { "Header2", "PropertyName2" },
                                                                     },
                                                                     new List<Object>()
                                                                     );
            return File(fileInfo.Contents, "application/vnd.ms-excel", fileInfo.Name);
        }
    }
}
