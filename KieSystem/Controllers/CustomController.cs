using KieSystem.DTOs;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;

namespace KieSystem.Controllers
{

    [Obsolete("This controller is deprecated and should not be used.")]
    public class ObsoleteControllerBase : ControllerBase
    {
    }

    [ApiController]
    [Route("api/[controller]")]
    public class CustomController : ObsoleteControllerBase
    {


        [HttpGet(Name = "get")]
        public IEnumerable<BlogClass> Get() { 
        
            IEnumerable<BlogClass> list = autogenerate(100);
            return list;
        }

        private List<BlogClass> autogenerate(int number) {
            List<BlogClass> list = new List<BlogClass>();
            for (int i = 0; i < number; i++) {
                list.Add(new BlogClass
                {
                    Id = i,
                    Title = "Title " + i.ToString(),
                    Body = "BOdy " + i.ToString(),
                    UserId = 1,
                });
            }
            return list;
        }
        [HttpGet("export",Name = "export")]
        public IActionResult ExportExcel()
        {
            var data = autogenerate(188);
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheetTaolao1");
            
            // Header
            IRow titleRow = sheet.CreateRow(0);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue("Title đầu tiên");
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0,0,0,3);
            sheet.AddMergedRegion(cellRangeAddress);

            //Column Title
            IRow headerRow = sheet.CreateRow(titleRow.RowNum + 1);


            //Row-Cell

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                var content = stream.ToArray();

                var result = new FileContentResult(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = "Demo.xlsx"
                };

                return result;
            }
        }

    }



}
