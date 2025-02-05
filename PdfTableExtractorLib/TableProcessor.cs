namespace PdfTableExtractorLib
{
    public static class TableProcessor
    {
        public static int CalculateRemovableColumnCount(string[][] data, RowWalkerDelegate walkerFunction)
        {
            return data.Select(columnData => CountRemovableColumnsInRow(walkerFunction, columnData))
                       .DefaultIfEmpty(0)
                       .Min();
        }

        private static int CountRemovableColumnsInRow(RowWalkerDelegate walkerFunction, string[] columnData)
        {
            return columnData
                .Select((_, index) => walkerFunction(columnData, index))
                .TakeWhile(string.IsNullOrWhiteSpace)
                .Count();
        }

        public static int CalculateRemovableRowCount(string[][] data, RowProviderDelegate rowProvider)
        {
            return data
                .Select((_, index) => rowProvider(data, index))
                .TakeWhile(row => row.All(string.IsNullOrWhiteSpace))
                .Count();
        }
    }
}
