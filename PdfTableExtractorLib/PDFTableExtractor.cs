using NPOI.XSSF.UserModel;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;

namespace PdfTableExtractorLib
{
    public static class PDFTableExtractor
    {
        public const int SKIP_METHOD_NONE = 0;
        public const int SKIP_METHOD_LEADING = 1;
        public const int SKIP_METHOD_TRAILING = 2;

        public delegate bool ComparisonDelegate(int value);

        public static XSSFWorkbook Extract(
            bool autosizeColumns,
            int emptyColumnSkipMethod,
            int emptyRowSkipMethod,
            Predicate<int> rowComparisonFunction,
            Predicate<int> columnComparisonFunction,
            PageNamingFunction pageNamingFunction,
            string pdfFilePath)
        {
            using var input = PdfDocument.Open(pdfFilePath);
            return DoExtract(autosizeColumns, emptyColumnSkipMethod, emptyRowSkipMethod,
                rowComparisonFunction, columnComparisonFunction, pageNamingFunction, input);
        }

        public static XSSFWorkbook Extract(
            bool autosizeColumns,
            int emptyColumnSkipMethod,
            int emptyRowSkipMethod,
            Predicate<int> rowComparisonFunction,
            Predicate<int> columnComparisonFunction,
            PageNamingFunction pageNamingFunction,
            byte[] pdfData)
        {
            using var stream = new MemoryStream(pdfData);
            using var input = PdfDocument.Open(stream);
            return DoExtract(autosizeColumns, emptyColumnSkipMethod, emptyRowSkipMethod,
                rowComparisonFunction, columnComparisonFunction, pageNamingFunction, input);
        }

        private static XSSFWorkbook DoExtract(
            bool autosizeColumns,
            int emptyColumnSkipMethod,
            int emptyRowSkipMethod,
            Predicate<int> rowComparisonFunction,
            Predicate<int> columnComparisonFunction,
            PageNamingFunction pageNamingFunction,
            PdfDocument pdfInput)
        {
            var textExtractor = new SpreadsheetExtractionAlgorithm();
            var rawPages = ObjectExtractor.Extract(pdfInput);

            var excelOutput = new XSSFWorkbook();

            int pageIndex = 0;
            while (rawPages.MoveNext()) // Iterating through the PageIterator
            {
                var page = rawPages.Current; // Get the current PageArea
                var asd = page.GetRulings();
                var tables = textExtractor.Extract(page); // Extract tables from the page
                foreach (var table in tables) // Iterate through extracted tables
                {
                    WriteTable(autosizeColumns, emptyColumnSkipMethod, emptyRowSkipMethod,
                        rowComparisonFunction, columnComparisonFunction, pageNamingFunction,
                        excelOutput, pageIndex, table);
                }
                pageIndex++;
            }

            return excelOutput;
        }

        private static void WriteTable(
            bool autosizeColumns,
            int emptyColumnSkipMethod,
            int emptyRowSkipMethod,
            Predicate<int> rowComparisonFunction,
            Predicate<int> columnComparisonFunction,
            PageNamingFunction pageNamingFunction,
            XSSFWorkbook excelOutput,
            int pageIndex,
            Table table)
        {
            var rawRows = table.Rows;

            if (rawRows.Count > 0)
            {
                var rows = rawRows.Select(row => row.Select(cell => cell.GetText()).ToArray()).ToArray();

                var removableColumnCountFromBegin = (emptyColumnSkipMethod & 1) == 0 ? 0 : TableProcessor.CalculateRemovableColumnCount(rows, RowWalkerFunction.WalkForwards());
                var removableColumnCountFromEnd = (emptyColumnSkipMethod & 2) == 0 ? 0 : TableProcessor.CalculateRemovableColumnCount(rows, RowWalkerFunction.WalkBackwards());
                var removableRowCountFromBegin = (emptyRowSkipMethod & 1) == 0 ? 0 : TableProcessor.CalculateRemovableRowCount(rows, RowProviderFunction.ProvidingForwards());
                var removableRowCountFromEnd = (emptyRowSkipMethod & 2) == 0 ? 0 : TableProcessor.CalculateRemovableRowCount(rows, RowProviderFunction.ProvidingBackwards());

                if (rowComparisonFunction(rows.Length - removableRowCountFromBegin - removableRowCountFromEnd) &&
                    columnComparisonFunction(rows[0].Length - removableColumnCountFromBegin - removableColumnCountFromEnd))
                {
                    var pageSheet = excelOutput.CreateSheet(
                        pageNamingFunction(excelOutput, pageIndex));

                    for (int rowIndex = removableRowCountFromBegin;
                         rowIndex < rows.Length - removableRowCountFromEnd;
                         rowIndex++)
                    {
                        WriteRow(rows, removableColumnCountFromBegin, removableColumnCountFromEnd,
                            pageSheet, rowIndex);
                    }

                    if (autosizeColumns)
                    {
                        for (int i = 0; i < pageSheet.GetRow(0).PhysicalNumberOfCells; i++)
                        {
                            pageSheet.AutoSizeColumn(i);
                        }
                    }

                    //pageSheet.SetActiveCellRange(1,1,1,1);
                }
            }
        }

        private static void WriteRow(
            string[][] rows,
            int removableColumnCountFromBegin,
            int removableColumnCountFromEnd,
            NPOI.SS.UserModel.ISheet pageSheet,
            int rowIndex)
        {
            var excelRow = pageSheet.CreateRow(rowIndex);

            for (int columnIndex = removableColumnCountFromBegin;
                 columnIndex < rows[rowIndex].Length - removableColumnCountFromEnd;
                 columnIndex++)
            {
                excelRow.CreateCell(columnIndex).SetCellValue(rows[rowIndex][columnIndex]);
            }
        }


        public static PageNamingFunction GetPageNamingFunction(int namingMethod)
        {
            return namingMethod == 0
                ? (workbook, pageIndex) => $"{workbook.NumberOfSheets + 1}."
                : (workbook, pageIndex) => $"{pageIndex + 1}. Table";
        }

        public static Predicate<int> GetComparisonFunction(int comparisonMethod, int value)
        {
            return comparisonMethod switch
            {
                0 => k => k <= value,
                1 => k => k == value,
                _ => k => k >= value
            };
        }
    }
}