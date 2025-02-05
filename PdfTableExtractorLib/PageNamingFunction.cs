using NPOI.XSSF.UserModel;

namespace PdfTableExtractorLib
{
    public delegate string PageNamingFunction(XSSFWorkbook workbook, int pageIndex);
}