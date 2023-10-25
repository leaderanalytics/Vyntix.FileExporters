namespace LeaderAnalytics.Vyntix.FileExporters;

public class MatrixFileExporter : IFileExporter
{
    public AsyncResult<string[,]> ToArray(FileExportArgs args, List<Vintage> vintages)
    {
        // The reason this method takes a list of Vintages versus a list of observations is that we want 
        // to include all Vintages for a series including those that have no observations.

        AsyncResult<string[,]> result = new AsyncResult<string[,]>();

        if (!(vintages?.Any() ?? false))
            return result;

        List<Observation> observations = vintages.SelectMany(x => x.Observations).ToList();

        if (observations.GroupBy(x => new { x.Vintage.NativeID, x.Vintage.SeriesDataProvider, x.Vintage.DataProviderID }).Count() > 1)
        {
            result.ErrorMessage = "Vintages can be for a single NativeID and from a single Data Provider only.";
            return result;
        }

        if (observations.GroupBy(x => x.ObsModelReference).Count() > 1)
        {
            result.ErrorMessage = "All Observations must have the same Model Reference.";
            return result;
        }

        // Create date to column index mappings.  Add 1 because column 0 and row 0 are headers.
        List<DateTime> columnsTmp = null;  
        
        if(args.VintSortDirection == SortDirection.Ascending)
            columnsTmp = observations.Select(x => x.Vintage.VintageDate).OrderBy(x => x).Distinct().ToList();
        else
            columnsTmp = observations.Select(x => x.Vintage.VintageDate).OrderByDescending(x => x).Distinct().ToList();

        List<DateTime> rowsTmp = null;
        
        if(args.VintSortDirection == SortDirection.Ascending)
            rowsTmp = observations.Select(x => x.ObsDate).OrderBy(x => x).Distinct().ToList();
        else
            rowsTmp = observations.Select(x => x.ObsDate).OrderByDescending(x => x).Distinct().ToList();

        Dictionary<DateTime, int> columns = columnsTmp.ToDictionary(x => x, y => columnsTmp.IndexOf(y) + 1);
        Dictionary<DateTime, int> rows = rowsTmp.ToDictionary(x => x, y => rowsTmp.IndexOf(y) + 1);

        // Create the matrix.  Add one row and one column for headers
        string[,] matrix = new string[observations.Count + 1, columns.Count + 1];
        // create row headers
        columns.Keys.ToList().ForEach(x => matrix[0, columns[x]] = x.ToString(args.DateFormat));

        for (int i = 0; i < observations.Count; i++)
        {
            Observation o = observations[i];
            int rowIndex = rows[o.ObsDate];
            int colIndex = columns[o.Vintage.VintageDate];
            matrix[rowIndex, 0] = o.ObsDate.ToString(args.DateFormat);
            matrix[rowIndex, colIndex] = o.Close.ToString();
        }

        result.Result = matrix;
        result.Success = true;
        return result;
    }

    public AsyncResult<string> ToCSV(FileExportArgs args, List<Vintage> vintages)
    {
        AsyncResult<string> result = new AsyncResult<string>();
        AsyncResult<string[,]> matrixResult = ToArray(args, vintages);

        if (!matrixResult.Success)
        {
            result.ErrorMessage = matrixResult.ErrorMessage;
            return result;
        }

        string[,] obsMatrix = matrixResult.Result;

        if (obsMatrix == null)
            return result;

        int rowCount = obsMatrix.GetLength(0);
        int colCount = obsMatrix.GetLength(1);
        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < colCount; j++)
                sb.Append(obsMatrix[i, j] + (j == colCount - 1 ? null : ","));

            sb.AppendLine();
        }
        result.Result = sb.ToString();
        result.Success = true;
        return result;
    }

    public AsyncResult<byte[]> ToExcel(FileExportArgs args, List<Vintage> vintages) 
    {
        AsyncResult<byte[]> result = new();
        AsyncResult<string[,]> matrixResult = ToArray(args, vintages);
        
        if(!matrixResult.Success) 
        {
            result.ErrorMessage = matrixResult.ErrorMessage;
            return result;
        }
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add(vintages.First().NativeID);
        string[,] m = matrixResult.Result;
        
        for (int r = 0; r < m.GetLength(0); r++)
        {
            for (int c = 0; c < m.GetLength(1); c++)
            {

                if (r == 0 ^ c == 0)
                {
                    ws.Cell(r + 1, c + 1).Value = Convert.ToDateTime(m[r, c]);
                    ws.Cell(r + 1, c + 1).Style.DateFormat.Format = "yyyy-MM-dd";
                }
                else
                    ws.Cell(r + 1, c + 1).Value = m[r, c];
            }
        }
        ws.Columns().AdjustToContents();

        using (MemoryStream ms = new())
        {
            try
            {
                wb.SaveAs(ms);
                result.Result = ms.ToArray();
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.ToString();
            }
            return result;
        }
    }
}
