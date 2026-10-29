using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Centrum.Pages;

public class SecureScoreModel(GraphServiceClient graph, ILogger<SecureScoreModel> logger) : PageModel
{
    public Microsoft.Graph.Models.SecureScore? Latest { get; private set; }
    public List<Microsoft.Graph.Models.SecureScore> History { get; private set; } = [];
    public Dictionary<string, SecureScoreControlProfile> Profiles { get; private set; } = [];
    public string? Error { get; private set; }

    public double? Percentage =>
        Latest?.CurrentScore is double c && Latest.MaxScore is double m and > 0 ? c / m * 100 : null;

    public async Task OnGetAsync(CancellationToken ct)
    {
        try
        {
            // Graph devuelve los resultados diarios, el más reciente primero.
            var scores = await graph.Security.SecureScores.GetAsync(r => r.QueryParameters.Top = 30, ct);
            History = scores?.Value ?? [];
            Latest = History.FirstOrDefault();

            var profiles = await graph.Security.SecureScoreControlProfiles.GetAsync(r => r.QueryParameters.Top = 999, ct);
            if (profiles is not null)
            {
                var iterator = PageIterator<SecureScoreControlProfile, SecureScoreControlProfileCollectionResponse>
                    .CreatePageIterator(graph, profiles, p =>
                    {
                        if (p.Id is not null) Profiles[p.Id] = p;
                        return true;
                    });
                await iterator.IterateAsync(ct);
            }
        }
        catch (ODataError ex)
        {
            logger.LogError(ex, "Graph error reading secure score");
            Error = $"{ex.Error?.Code}: {ex.Error?.Message}";
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error reading secure score");
            Error = ex.Message;
        }
    }
}
