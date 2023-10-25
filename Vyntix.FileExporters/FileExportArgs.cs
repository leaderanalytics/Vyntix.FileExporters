namespace LeaderAnalytics.Vyntix.FileExporters;


public enum DataLayout
{
    List,
    Matrix
}

public enum FileFormat
{
    CSV,
    Excel
}

public enum SortPriority
{
    ObservationDate,
    VintageDate
}

public enum SortDirection
{
    Ascending,
    Descending
}

public class FileExportArgs
{
    public DataLayout DataLayout { get; set; }
    public FileFormat FileFormat { get; set; }
    public SortPriority SortPriority { get; set; }
    public SortDirection ObsSortDirection { get; set; }
    public SortDirection VintSortDirection { get; set; }
    public string DateFormat { get; set; } = "yyyy-MM-dd";
}