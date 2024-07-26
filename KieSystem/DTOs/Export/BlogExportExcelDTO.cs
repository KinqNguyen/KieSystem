using NPOI.SS.Formula.Functions;

namespace KieSystem.DTOs
{
    public class ExportExcelDTO
    {
        public GeneralFormat generalInfo { get; set; } = new ExportExcelDTO();
        public List<ExcelColumn> columns { get; set; } = new List<ExcelColumn>();
        public List<ExcelRow<T>> Rows { get; set; } = new List<ExcelRow<T>> { };    
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
        public bool isAlternate { get; set; }
        public GeneralFormat()
        {
            isAlternate = false;
        }

    }

}
