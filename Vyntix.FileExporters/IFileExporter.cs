using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeaderAnalytics.Vyntix.FileExporters;
public interface IFileExporter
{
    AsyncResult<string> ToCSV(FileExportArgs args, List<Vintage> vintages);
    AsyncResult<byte[]> ToExcel(FileExportArgs args, List<Vintage> vintages);
}
