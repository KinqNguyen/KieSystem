using KieSystem.DTOs;
using KieSystem.Interface;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Reflection;



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
        private readonly IExcelService _excelService;

        public CustomController(IExcelService excelExportService)
        {
            _excelService = excelExportService;
        }




        [HttpGet(Name = "get")]
        public IEnumerable<BlogDTO> Get()
        {

            IEnumerable<BlogDTO> list = autogenerate(100);
            return list;
        }



        private List<BlogDTO> autogenerate(int number)
        {
            List<BlogDTO> list = new List<BlogDTO>();
            int totalViewBase = 1000;
            int totalViewIncrement = 1000;
            DateTime oneMonthAgo = DateTime.Now.AddMonths(-1);

            for (int i = 0; i < number; i++)
            {
                list.Add(new BlogDTO
                {
                    Id = i,
                    Title = "Title " + i.ToString(),
                    Body = "Body " + i.ToString(),

                    UserId = 100 + i,
                    Totalview = totalViewBase + (i % 10) * totalViewIncrement, // TotalView from 1000 to 10000
                    Releasedate = oneMonthAgo.AddDays(i), // ReleaseDate one month ago plus i days
                    Userlevel = (i % 4) + 1 // UserLevel from 1 to 4
                });
            }
            return list;
        }




        [HttpPost("export", Name = "export")]
        public IActionResult ExportExcel([FromBody] ExportExcelDTO exportDto)
        {
            try
            {
                // Kiểm tra xem dữ liệu có hợp lệ không
                if (exportDto == null || !ModelState.IsValid)
                {
                    return BadRequest("Invalid data.");
                }

                var data = autogenerate(188); // Giả định có hàm autogenerate để tạo dữ liệu
                byte[] content = _excelService.ExportExcel(exportDto, data);

                return new FileContentResult(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = "Kiet.xlsx"
                };
            }
            catch (Exception ex)
            {
                // Log lỗi và trả về mã lỗi 400 nếu có lỗi dữ liệu đầu vào
                return BadRequest("An error occurred while processing your request.");
            }
        }

        [HttpPost("exportsimple", Name = "exportsimple")]
        public IActionResult ExportSimpleExcel([FromBody] SimpleExportDto exportDto)
        {
            try
            {

                var data = autogenerate(188); // Giả định có hàm autogenerate để tạo dữ liệu
                byte[] content = _excelService.RawExportExcel(exportDto.ColumnExport, exportDto.HeaderTitle, data);

                return new FileContentResult(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = "Kiet.xlsx"
                };
            }
            catch (Exception ex)
            {
                // Log lỗi và trả về mã lỗi 400 nếu có lỗi dữ liệu đầu vào
                return BadRequest("An error occurred while processing your request.");
            }
        }


        [HttpGet("create-demo-excel")]
        public IActionResult CreateDemoExcel()
        {
            // Tạo một workbook mới
            IWorkbook workbook = new XSSFWorkbook();

            // Tạo một sheet mới trong workbook
            ISheet sheet = workbook.CreateSheet("Holy Sheet1");

            //// Tạo một hàng mới (hàng thứ 0)
            IRow row = sheet.CreateRow(0);

            // Tạo một cell mới trong hàng đầu tiên (cell đầu tiên, chỉ số 0)
            ICell cell = row.CreateCell(0);
            // Đặt giá trị cho cell
            cell.SetCellValue("Hello, Kiet!");

            //// Áp dụng định dạng cho cell
            //ICellStyle cellStyle = workbook.CreateCellStyle();
            //IFont font = workbook.CreateFont();
            //font.IsBold = true;
            //cellStyle.SetFont(font);
            //cell.CellStyle = cellStyle;

            //// Tạo một cell khác trong cùng hàng và đặt giá trị là một số
            //ICell numericCell = row.CreateCell(1);
            //numericCell.SetCellValue(123);

            //// Thêm một cell với công thức
            //ICell formulaCell = row.CreateCell(2);
            //formulaCell.CellFormula = "A1 & \" World\""; // Công thức nối chuỗi từ cell A1

            //// Merge cells (gộp các cell lại)
            //CellRangeAddress mergeRegion = new CellRangeAddress(0, 0, 3, 5); // Gộp các cell từ D1 đến F1
            //sheet.AddMergedRegion(mergeRegion);
            //ICell mergedCell = row.CreateCell(3);
            //mergedCell.SetCellValue("Merged Cells");

            //// Đặt màu nền cho cell (tìm màu gần nhất)
            //cellStyle.FillForegroundColor = IndexedColors.LightYellow.Index;
            //cellStyle.FillPattern = FillPattern.SolidForeground;
            //mergedCell.CellStyle = cellStyle;

            //// Tạo một cell khác và đặt công thức (tính toán đơn giản)
            //ICell calcCell = row.CreateCell(6);
            //calcCell.CellFormula = "B1 * 2"; // Công thức nhân đôi giá trị của cell B1

            // Ghi workbook ra file và trả về mảng byte
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                var byteArray = stream.ToArray();
                return File(byteArray, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Demo.xlsx");
            }
        }


    }



}
