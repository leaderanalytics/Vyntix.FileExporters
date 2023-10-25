namespace Vyntix.FileExporters.Tests;

public class MatrixTests : BaseTest
{
    [Test]
    public async Task SortAscendingTest()
    {
        FileExportArgs args = new() { VintSortDirection = SortDirection.Ascending, ObsSortDirection = SortDirection.Ascending};
        AsyncResult<string[,]> result = new MatrixFileExporter().ToArray(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        Assert.AreEqual(new DateTime(1960,2,1).ToString(args.DateFormat), result.Result[0, 1]);
        Assert.AreEqual(new DateTime(1960, 1, 1).ToString(args.DateFormat), result.Result[1, 0]);
    }

    [Test]
    public async Task SortDescendingTest()
    {
        FileExportArgs args = new() { VintSortDirection = SortDirection.Descending, ObsSortDirection = SortDirection.Descending };
        AsyncResult<string[,]> result = new MatrixFileExporter().ToArray(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        Assert.AreEqual(new DateTime(1960, 4, 1).ToString(args.DateFormat), result.Result[0, 1]);
        Assert.AreEqual(new DateTime(1960, 3, 1).ToString(args.DateFormat), result.Result[1, 0]);
    }


    [Test]
    public async Task ExcelTest()
    {
        FileExportArgs args = new() { VintSortDirection = SortDirection.Descending, ObsSortDirection = SortDirection.Descending };
        AsyncResult<byte[]> result = new MatrixFileExporter().ToExcel(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        System.IO.File.WriteAllBytes("MatrixTest.xlsx", result.Result);
    }
}