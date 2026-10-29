using System.ComponentModel.DataAnnotations;

namespace Centrum;

public class AzureAdOptions
{
    [Required] public string TenantId { get; set; } = "";
    [Required] public string ClientId { get; set; } = "";
    [Required] public string ClientSecret { get; set; } = "";
}
