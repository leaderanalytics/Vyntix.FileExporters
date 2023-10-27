using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeaderAnalytics.Vyntix.FileExporters;
public class FileExporter : IFileExporter
{
    public AsyncResult<byte[]> Export(FileExportArgs args, List<Vintage> vintages)
    {
        ArgumentNullException.ThrowIfNull(vintages);

        if (args.DataLayout == DataLayout.List)
        {
            ListFileExporter exporter = new ListFileExporter();

            if (args.FileFormat == FileFormat.Excel)
                return exporter.ToExcel(args, vintages);
            else
                return exporter.ToCSV(args, vintages);
        }
        else
        {
            MatrixFileExporter exporter = new MatrixFileExporter();

            if (args.FileFormat == FileFormat.Excel)
                return exporter.ToExcel(args, vintages);
            else
                return exporter.ToCSV(args, vintages);
        }
    }
}
