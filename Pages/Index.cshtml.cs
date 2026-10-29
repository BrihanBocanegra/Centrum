using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Centrum.Pages;

public class IndexModel(GraphServiceClient graph, ILogger<IndexModel> logger) : PageModel
{
    public List<User> Users { get; private set; } = [];
    public string? Error { get; private set; }

    public async Task OnGetAsync(CancellationToken ct)
    {
        try
        {
            var page = await graph.Users.GetAsync(r =>
            {
                r.QueryParameters.Select = ["id", "displayName", "mail", "userPrincipalName", "jobTitle", "accountEnabled"];
                r.QueryParameters.Top = 999;
                r.QueryParameters.Orderby = ["displayName"];
                r.QueryParameters.Count = true;
                r.Headers.Add("ConsistencyLevel", "eventual");
            }, ct);

            if (page is null) return;

            var iterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
                graph, page, u => { Users.Add(u); return true; },
                req => { req.Headers.Add("ConsistencyLevel", "eventual"); return req; });
            await iterator.IterateAsync(ct);
        }
        catch (ODataError ex)
        {
            logger.LogError(ex, "Graph error listing users");
            Error = $"{ex.Error?.Code}: {ex.Error?.Message}";
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Error listing users");
            Error = ex.Message;
        }
    }
}
