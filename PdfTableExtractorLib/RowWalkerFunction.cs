namespace PdfTableExtractorLib
{
    public delegate string RowWalkerDelegate(string[] row, int columnIndex);

    public static class RowWalkerFunction
    {
        public static RowWalkerDelegate WalkForwards()
        {
            return (row, columnIndex) => row[columnIndex];
        }

        public static RowWalkerDelegate WalkBackwards()
        {
            return (row, columnIndex) => row[row.Length - 1 - columnIndex];
        }
    }
}