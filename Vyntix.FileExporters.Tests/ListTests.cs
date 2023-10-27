using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Vyntix.FileExporters.Tests;
public class ListTests : BaseTest
{
    [Test]
    public async Task SortAscendingTest()
    {
        FileExportArgs args = new() { VintSortDirection = SortDirection.Ascending, ObsSortDirection = SortDirection.Ascending };
        AsyncResult<byte[]> result = new ListFileExporter().ToCSV(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        
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
        AsyncResult<byte[]> result = new ListFileExporter().ToExcel(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        System.IO.File.WriteAllBytes("ListTest.xlsx", result.Result);
    }

    [Test]
    public async Task CsvTest()
    {
        FileExportArgs args = new() { VintSortDirection = SortDirection.Descending, ObsSortDirection = SortDirection.Descending };
        AsyncResult<byte[]> result = new MatrixFileExporter().ToCSV(args, vintages);
        Assert.IsTrue(result.Success);
        Assert.IsNotNull(result.Result);
        System.IO.File.WriteAllBytes("MatrixTest.csv", result.Result);
    }
}
