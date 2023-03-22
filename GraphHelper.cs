using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Me;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        return _userClient.Me
            .GetAsync((requestConfiguration) => 
            {
                // Only request specific properties
                requestConfiguration.QueryParameters.Select = new [] { "displayName", "mail", "userPrincipalName" };
            });
    }
    // This function serves as a playground for testing Graph snippets
    // or other code
    public async static Task<ApplicationCollectionResponse?> ListApplicationsAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        var applications = await _userClient.Applications.GetAsync();
        return applications;
   }
    public async static Task<ServicePrincipalCollectionResponse?> ListServicePrincipalsAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Top = 900;
        });
        return servicePrincipals;
    }
    public async static Task<ServicePrincipalCollectionResponse?> ListServicePrincipalsWithoutMicrosoftAppsAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Filter = $"appOwnerOrganizationId ne {_settings.MicrosoftAppTenantId}"; 
            config.QueryParameters.Count = true;
            config.Headers.Add("ConsistencyLevel", "eventual");
            config.QueryParameters.Top = 900;
        });
        return servicePrincipals;
   }
    public async static Task<ServicePrincipalCollectionResponse?> ListServicePrincipalsEnterpriseAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Filter = $"tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')"; 
            config.QueryParameters.Count = true;
            config.Headers.Add("ConsistencyLevel", "eventual");
            config.QueryParameters.Top = 900;
        });
        return servicePrincipals;
   }
    public async static Task<ServicePrincipalCollectionResponse?> ListManagedIdentitiesAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Filter = $"servicePrincipalType eq 'ManagedIdentity'"; 
            config.QueryParameters.Count = true;
            config.Headers.Add("ConsistencyLevel", "eventual");
            config.QueryParameters.Top = 900;
        });
        return servicePrincipals;
   }
   public async static Task<ServicePrincipalCollectionResponse?> GetApplicatonPermissionsAsync(string appName)
   {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var spn = await _userClient.ServicePrincipals.GetAsync((config) =>
        {
            config.QueryParameters.Filter = $"startsWith(displayName,'{appName}')";
        }
        );
        return spn;
   }
}
