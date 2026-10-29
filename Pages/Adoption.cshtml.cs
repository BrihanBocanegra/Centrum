using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace Centrum.Pages;

public record ServiceAdoption(string Name, int Active, int Inactive)
{
    public int Total => Active + Inactive;
    public double Percentage => Total > 0 ? (double)Active / Total * 100 : 0;
}

public class AdoptionModel(GraphServiceClient graph, ILogger<AdoptionModel> logger) : PageModel
{
    public static readonly string[] Periods = ["D7", "D30", "D90", "D180"];

    [BindProperty(SupportsGet = true)]
    public string Period { get; set; } = "D30";

    public List<ServiceAdoption> Services { get; private set; } = [];
    public string? RefreshDate { get; private set; }
    public string? Error { get; private set; }

    public async Task OnGetAsync(CancellationToken ct)
    {
        if (!Periods.Contains(Period)) Period = "D30";

        try
        {
            using var stream = await graph.Reports
                .GetOffice365ServicesUserCountsWithPeriod(Period)
                .GetAsync(cancellationToken: ct);
            if (stream is null) return;

            using var reader = new StreamReader(stream);
            var header = (await reader.ReadLineAsync(ct))?.TrimStart('\uFEFF').Split(',');
            var row = (await reader.ReadLineAsync(ct))?.Split(',');
            if (header is null || row is null || header.Length != row.Length) return;

            var values = header.Zip(row).ToDictionary(p => p.First.Trim(), p => p.Second.Trim());
            values.TryGetValue("Report Refresh Date", out var date);
            RefreshDate = date;

            foreach (var name in new[] { "Exchange", "OneDrive", "SharePoint", "Teams", "Yammer", "Skype For Business", "Office 365" })
            {
                if (int.TryParse(values.GetValueOrDefault($"{name} Active"), out var active) &&
                    int.TryParse(values.GetValueOrDefault($"{name} Inactive"), out var inactive))
                {
                    Services.Add(new ServiceAdoption(name == "Office 365" ? "Microsoft 365 (global)" : name, active, inactive));
                }
            }
        }
        catch (ODataError ex)
        {
            logger.LogError(ex, "Graph error reading adoption report");
            Error = $"{ex.Error?.Code}: {ex.Error?.Message}";
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error reading adoption report");
            Error = ex.Message;
        }
    }
}
