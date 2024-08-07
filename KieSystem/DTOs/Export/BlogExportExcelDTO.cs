using NPOI.SS.Formula.Functions;

namespace KieSystem.DTOs
{
    public class ExportExcelDTO
    {
        public GeneralFormat generalInfo { get; set; } = new GeneralFormat();

    }

    public class SimpleExportDto
    {
        public List<string> ColumnExport { get; set; }
        public string HeaderTitle { get; set; }
    }

    public class ExcelRow<T>
    {
        public Format format { get; set; } = new Format();
        public T rowData { get; set; }  
    }
    public class ExcelColumn
    {
        public Format format { get; set; } = new Format();  
        public string title { get; set;} = "";
    }
    public class Format
    {
        public string backgroundColor { get; set; }
    }

    public class GeneralFormat
    {
        public string FileName { get; set; }
        public HeaderTitle HeaderTitle { get; set; }
        public Alternate Alternate { get; set; }
        public int PriorityColor { get; set; }
        public List<string> ColumnExport { get; set; }

    }

    public class HeaderTitle
    {
        public string HeaderName { get; set; }
        public Font Font { get; set; }
        public string Color { get; set; }
        public Merge Merge { get; set; }
    }

    public class Font
    {
        public bool Bold { get; set; }
        public bool Underline { get; set; }
        public bool Italic { get; set; }
    }

    public class Merge
    {
        public bool Enable { get; set; }
        public int Count { get; set; }
    }

    public class Alternate
    {
        public bool Enable { get; set; }
        public string Color { get; set; }
    }

}
