namespace PdfTableExtractorLib
{
    public delegate string[] RowProviderDelegate(string[][] table, int rowIndex);

    public static class RowProviderFunction
    {
        public static RowProviderDelegate ProvidingForwards()
        {
            return (table, rowIndex) => table[rowIndex];
        }

        public static RowProviderDelegate ProvidingBackwards()
        {
            return (table, rowIndex) => table[table.Length - 1 - rowIndex];
        }
    }
}