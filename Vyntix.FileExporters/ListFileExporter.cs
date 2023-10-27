using CsvHelper.Configuration;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace LeaderAnalytics.Vyntix.FileExporters;

public class ListFileExporter
{
    private IEnumerable<CSVVintatage> BuildVintages(FileExportArgs args, List<Vintage> vintages)
    {
        IEnumerable<CSVVintatage> observations = vintages.SelectMany(x => x.Observations).Select(x => new CSVVintatage(x, args.DateFormat));

        if (args.SortPriority == SortPriority.ObservationDate)
        {
            if (args.ObsSortDirection == SortDirection.Ascending)
            {
                if (args.VintSortDirection == SortDirection.Ascending)
                    observations = observations.OrderBy(x => x.ObsDate).ThenBy(x => x.VintageDate);
                else
                    observations = observations.OrderBy(x => x.ObsDate).ThenByDescending(x => x.VintageDate);
            }
            else
            {
                if (args.VintSortDirection == SortDirection.Ascending)
                    observations = observations.OrderByDescending(x => x.ObsDate).ThenBy(x => x.VintageDate);
                else
                    observations = observations.OrderByDescending(x => x.ObsDate).ThenByDescending(x => x.VintageDate);
            }
        }
        else
        {
            if (args.VintSortDirection == SortDirection.Ascending)
            {
                if (args.ObsSortDirection == SortDirection.Ascending)
                    observations = observations.OrderBy(x => x.VintageDate).ThenBy(x => x.ObsDate);
                else
                    observations = observations.OrderBy(x => x.VintageDate).ThenByDescending(x => x.ObsDate);
            }
            else
            {
                if (args.ObsSortDirection == SortDirection.Ascending)
                    observations = observations.OrderByDescending(x => x.VintageDate).ThenBy(x => x.ObsDate);
                else
                    observations = observations.OrderByDescending(x => x.VintageDate).ThenByDescending(x => x.ObsDate);
            }
        }
        return observations;
    }

    public AsyncResult<byte[]> ToCSV(FileExportArgs args, List<Vintage> vintages)
    {
        AsyncResult<byte[]> result = new();

        if (!(vintages?.Any() ?? false))
            return result;

        IEnumerable<CSVVintatage> observations = BuildVintages(args, vintages);

        using (var stream = new MemoryStream())
        {
            using (var writer = new StreamWriter(stream))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            csv.WriteRecords(observations);
            result.Result = stream.ToArray();
        }
        result.Success = true;
        return result;
    }

    public AsyncResult<byte[]> ToExcel(FileExportArgs args, List<Vintage> vintages)
    {
        AsyncResult<byte[]> result = new();

        if (!(vintages?.Any() ?? false))
            return result;

        List<CSVVintatage> csvVintages = BuildVintages(args, vintages).ToList();
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add(vintages.First().NativeID);

        ws.Cell(1, 1).Value = nameof(CSVVintatage.ObsDate);
        ws.Cell(1, 2).Value = nameof(CSVVintatage.VintageDate);
        ws.Cell(1, 3).Value = nameof(CSVVintatage.Open);
        ws.Cell(1, 4).Value = nameof(CSVVintatage.High);
        ws.Cell(1, 5).Value = nameof(CSVVintatage.Low);
        ws.Cell(1, 6).Value = nameof(CSVVintatage.Close);
        ws.Cell(1, 7).Value = nameof(CSVVintatage.VintageName);
        ws.Cell(1, 8).Value = nameof(CSVVintatage.DataProviderID);
        ws.Cell(1, 9).Value = nameof(CSVVintatage.ModelReferenceUserID);
        ws.Cell(1, 10).Value = nameof(CSVVintatage.ModelReferenceName);
        ws.Cell(1, 11).Value = nameof(CSVVintatage.ModelReferenceNotes);
        
        for (int i = 0; i < csvVintages.Count(); i++)
        {
            CSVVintatage v = csvVintages[i];
            int c = 0;
            
            ws.Cell(i + 2, ++c).Value = Convert.ToDateTime(v.ObsDate);
            ws.Cell(i + 2, c).Style.DateFormat.Format = args.DateFormat;
            ws.Cell(i + 2, ++c).Value = Convert.ToDateTime(v.VintageDate);
            ws.Cell(i + 2, c).Style.DateFormat.Format = args.DateFormat;
            ws.Cell(i + 2, ++c).Value = v.Open;
            ws.Cell(i + 2, ++c).Value = v.High;
            ws.Cell(i + 2, ++c).Value = v.Low;
            ws.Cell(i + 2, ++c).Value = v.Close;
            ws.Cell(i + 2, ++c).Value = v.VintageName;
            ws.Cell(i + 2, ++c).Value = v.DataProviderID;
            ws.Cell(i + 2, ++c).Value = v.ModelReferenceUserID;
            ws.Cell(i + 2, ++c).Value = v.ModelReferenceName;
            ws.Cell(i + 2, ++c).Value = v.ModelReferenceNotes;
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

internal class CSVVintatage
{
    public string ObsDate { get; set; }
    public string VintageDate { get; set; }
    public decimal Open { get; set; }
    public decimal High { get; set; }
    public decimal Low { get; set; }
    public decimal Close { get; set; }
    public string VintageName { get; set; }
    public int DataProviderID { get; set; }
    public int ModelReferenceUserID { get; set; }
    public string ModelReferenceName { get; set; }
    public string? ModelReferenceNotes { get; set; }

    public CSVVintatage(Observation o, string dateFormat)
    {
        ArgumentNullException.ThrowIfNull(o);

        ObsDate = o.ObsDate.ToString(dateFormat);
        Open = o.Open;
        High = o.High;
        Low = o.Low;
        Close = o.Close;
        VintageDate = o.Vintage.VintageDate.ToString(dateFormat);
        VintageName = o.Vintage.Name;
        DataProviderID = o.Vintage.DataProviderID;
        ModelReferenceUserID = o.ObsModelReference?.ModelReference.UserID ?? 0;
        ModelReferenceName = o.ObsModelReference?.ModelReference.Name;
        ModelReferenceNotes = o.ObsModelReference?.ModelReference.Notes;
    }
}