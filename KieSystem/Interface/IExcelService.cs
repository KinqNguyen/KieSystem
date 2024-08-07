using KieSystem.DTOs;

namespace KieSystem.Interface
{
    public interface IExcelService
    {
        byte[] ExportExcel(ExportExcelDTO exportDto, IEnumerable<BlogDTO> data);
        byte[] RawExportExcel(IEnumerable<string> columns, string HeaderTitle, IEnumerable<BlogDTO> data);
    }
}
