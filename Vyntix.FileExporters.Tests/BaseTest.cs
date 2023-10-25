using Microsoft.Extensions.Logging;
using Serilog;

namespace Vyntix.FileExporters.Tests;
public abstract class BaseTest
{
    protected List<Vintage> vintages;

    [SetUp]
    public virtual async Task Setup()
    {
        // Create some test data
        vintages = new();
        string symbol = "GNPCA";
        DateTime obsStartDate = new DateTime(1960, 1, 1);
        DateTime vintStartDate = new DateTime(1960, 2, 1);
        ObsModelReference omr = new ObsModelReference { ModelReference = new ModelReference { Name = "Test", Notes = "Notes" } };

        for (int i = 0; i < 3; i++)
        {
            Vintage v = new Vintage
            {
                VintageDate = vintStartDate.AddMonths(i),
                NativeID = symbol,
                DataProviderID = 1,
                SeriesDataProviderID = 1,
                Name = i.ToString(),
                
            };
            v.Observations = new List<Observation> { new Observation 
            { 
                ObsDate = obsStartDate.AddMonths(i), 
                Open = i -3,
                High = i - 2,
                Low = i - 1,
                Close = i, 
                Vintage = v,
                ObsModelReference = omr
            } };
            vintages.Add(v);
        }

        Assert.That(3, Is.EqualTo(vintages.Count));
    }
}
