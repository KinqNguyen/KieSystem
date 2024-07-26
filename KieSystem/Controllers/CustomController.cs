using KieSystem.DTOs;
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





        [HttpGet(Name = "get")]
        public IEnumerable<BlogDTO> Get()
        {

            IEnumerable<BlogDTO> list = autogenerate(100);
            return list;
        }



        private List<BlogDTO> autogenerate(int number)
        {
            List<BlogDTO> list = new List<BlogDTO>();
            for (int i = 0; i < number; i++)
            {
                list.Add(new BlogDTO
                {
                    Id = i,
                    Title = "Title " + i.ToString(),
                    Body = "BOdy " + i.ToString(),
                    UserId = 1,
                });
            }
            return list;
        }
        [HttpPost("export", Name = "export")]
        public IActionResult ExportExcel(ExportExcelDTO exportDto)
        {
            var data = autogenerate(188);
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("sheetTaolao1");

            // Header
            IRow titleRow = sheet.CreateRow(0);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue("Title đầu tiênnn");
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 3);
            sheet.AddMergedRegion(cellRangeAddress);



            //Column Title
            PropertyInfo[] props = typeof(BlogDTO).GetProperties();
            var headerRowNum = titleRow.RowNum + 1;
            IRow headerRow = sheet.CreateRow(headerRowNum);
            for (int i = 0; i < props.Length; i++)
            {
                // Tạo font và style cho cell
                ICellStyle cellStyle = workbook.CreateCellStyle();
                IFont font = workbook.CreateFont();

                // Định dạng underline, italic và bold cho font
                font.Underline = FontUnderlineType.Single;
                font.IsItalic = true;
                font.IsBold = true;
                cellStyle.SetFont(font);


                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(props[i].Name);
                cell.CellStyle = cellStyle;
            }



            //Row-Cell
            int firstDataRowIndex = headerRowNum+1;
            foreach (var item in data)
            {
                int currentIndex = firstDataRowIndex++;
                IRow row = sheet.CreateRow(currentIndex);
                if (exportDto.generalInfo.isAlternate) {
                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    IFont font = workbook.CreateFont();
                    if (currentIndex % 2==0) {
                        // Định dạng underline, italic và bold cho font
                        font.Underline = FontUnderlineType.Single;
                        font.IsItalic = true;
                        font.IsBold = true;
                        cellStyle.SetFont(font);
                        cellStyle.FillForegroundColor = IndexedColors.LightBlue.Index;
                        cellStyle.FillPattern = FillPattern.SolidForeground;
                    }
                    for (int i = 0; i < props.Length; i++)
                    {
                        ICell cell = row.CreateCell(i);
                        var value = props[i].GetValue(item)?.ToString() ?? string.Empty;
                        cell.SetCellValue(value);
                        cell.CellStyle = cellStyle;
                    }
                }
                else {
                    for (int i = 0; i < props.Length; i++)
                    {
                        ICell cell = row.CreateCell(i);
                        var value = props[i].GetValue(item)?.ToString() ?? string.Empty;
                        cell.SetCellValue(value);
                    }
                }
            }





            // Export file
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                var content = stream.ToArray();



                var result = new FileContentResult(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = "Demo2.xlsx"
                };



                return result;
            }
        }



    }





}
