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
    private static Dictionary<string, ServicePrincipal> _apis = new Dictionary<string, ServicePrincipal>();

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
    public async static Task<ApplicationCollectionResponse?> ListApplicationsAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        var applications = await _userClient.Applications.GetAsync();
        return applications;
    }
    public async static Task<IList<Application>> ListApplicationsWithoutServicePrincipalAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        var applications = await _userClient.Applications.GetAsync();
        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Top = 900;
        });
        List<Application> appList = new List<Application>();
        if (null != applications?.Value && null != servicePrincipals?.Value)
        {
            foreach ( var app in applications.Value )
            {
                // Check if SPN with appId is found
                var spn = servicePrincipals.Value.FirstOrDefault(spn => (spn.AppId == app.AppId));
                if (null == spn)
                {
                    appList.Add(app);
                }
            }
        }
        return appList;
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
    public async static Task<ServicePrincipalCollectionResponse?> ListServicePrincipalsWithInternalApplicationAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Filter = $"appOwnerOrganizationId eq {_settings.TenantId}"; 
            config.QueryParameters.Count = true;
            config.Headers.Add("ConsistencyLevel", "eventual");
            config.QueryParameters.Top = 900;
        });
        return servicePrincipals;
   }
    public async static Task<ServicePrincipalCollectionResponse?> ListServicePrincipalsWithExternalApplicationAsync()
    {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var servicePrincipals = await _userClient.ServicePrincipals.GetAsync((config) => 
        {
            config.QueryParameters.Filter = $"tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp') and appOwnerOrganizationId ne {_settings.TenantId}"; 
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
        if (null != spn?.Value)
        {
            foreach (var sp in spn.Value)
            {
                var permissions = await _userClient.ServicePrincipals[sp.Id].Oauth2PermissionGrants.GetAsync((config) =>
                {
                    config.QueryParameters.Top = 900;
                }
                );
                if (null != permissions?.Value)
                {
                    foreach (var permission in permissions.Value)
                    {
                        // var app = await _userClient.Applications[permission.ClientId].GetAsync();
                        //if (null != app)
                        //{
                        //    Console.WriteLine($"App: {app.DisplayName} has permission {permission.Scope}");
                        //}
                        ServicePrincipal? api = null;
                        if (!String.IsNullOrEmpty(permission.ResourceId))
                        {
                            if (!_apis.TryGetValue(permission.ResourceId, out api))
                            {
                                api = await _userClient.ServicePrincipals[permission.ResourceId].GetAsync();
                                if (null != api)
                                {
                                    _apis[permission.ResourceId] = api;
                                }
                            }
                        }
                        Console.WriteLine($"App: {sp.DisplayName} has permission {permission.Scope} | {permission.ConsentType} | {api?.AppDisplayName}");
                    }
                }
            }
        }
        return spn;
   }
   public async static Task<string> GetApplicatonPermissionsAsStringAsync(ServicePrincipal spn)
   {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

            var permissions = await _userClient.ServicePrincipals[spn.Id].Oauth2PermissionGrants.GetAsync((config) =>
            {
                config.QueryParameters.Top = 900;
            }
            );
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (null != permissions?.Value)
            {
                foreach (var permission in permissions.Value)
                {
                    ServicePrincipal? api = null;
                    if (!String.IsNullOrEmpty(permission.ResourceId))
                    {
                        if (!_apis.TryGetValue(permission.ResourceId, out api))
                        {
                            api = await _userClient.ServicePrincipals[permission.ResourceId].GetAsync();
                            if (null != api)
                            {
                                _apis[permission.ResourceId] = api;
                            }
                        }
                    }
                    sb.Append($"{api?.AppDisplayName}:{permission.Scope}-{permission.ConsentType}/");
                }
            }
        return sb.ToString();
   }
   public async static Task<string> GetAssignedUsersAsync(ServicePrincipal spn)
   {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");
            var appRoles = await _userClient.ServicePrincipals[spn.Id].AppRoleAssignedTo.GetAsync((config) =>
            {
                config.QueryParameters.Top = 900;
            }
            );
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            if (null != appRoles?.Value)
            {
                foreach (var appRole in appRoles.Value)
                {
                    
                    sb.Append($"{appRole.PrincipalDisplayName}-");
                }
            }
            return sb.ToString();
   }

   public async static Task<SignInCollectionResponse?> GetApplicatonSignInsAsync(string appName)
   {
         // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");
        _ = _settings ??
            throw new System.NullReferenceException("Settings not yet initialized.");

        var signInCollectionResponse = await _userClient.AuditLogs.SignIns.GetAsync((config) =>
        {
            config.QueryParameters.Filter = $"startsWith(appDisplayName,'{appName}')";
            config.QueryParameters.Top = 2;
        }
        );

        return signInCollectionResponse;
   }

   
}
