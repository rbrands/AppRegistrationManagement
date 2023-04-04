using System.Text;
using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

// See https://aka.ms/new-console-template for more information
// Created via tutorial https://learn.microsoft.com/en-us/graph/tutorials/dotnet?tabs=aad
Console.WriteLine(".NET Graph AppRegistrations\n");

Settings settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display access token");
    Console.WriteLine("2. List ServicePrincipals excluding Microsoft apps");
    Console.WriteLine("3. List ServicePrincipals");
    Console.WriteLine("4. List applications");
    Console.WriteLine("5. Get permissions requested by ServicePrincipal");
    Console.WriteLine("6. List ServicePrincipals of type EnterpriseApp");
    Console.WriteLine("7. List ManagedIdentities");
    Console.WriteLine("8. List ServicePrincipals refering an internal application");
    Console.WriteLine("9. List ServicePrincipals refering an external application");
    Console.WriteLine("10. List applications without ServicePrincipal");
    Console.WriteLine("11. List ServicePrincipals with verified publisher");
    Console.WriteLine("12. List Sign-Ins for application");


    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch(choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // Display access token
            await DisplayAccessTokenAsync();
            break;
        case 2:
            await ListServicePrincipalsAsync(true);
            break;
        case 3:
            // List ServicePrincipals
            await ListServicePrincipalsAsync();
            break;
        case 4:
            // List applications
            await ListApplicationsAsync();
            break;
        case 5:
            Console.WriteLine("Please enter an AppName to look for (emtpy for default):");
            string? appName = Console.ReadLine();
            if (String.IsNullOrWhiteSpace(appName))
            {
                appName = settings.AppDisplayName ?? "Azure";
            }
            await ListServicePrincipalPermissionsAsync(appName);
            break;
        case 6:
            await ListServicePrincipalsEnterpriseAsync();
            break;
        case 7:
            await ListManagedIdentitiesAsync();
            break;
        case 8:
            await ListServicePrincipalsWithInternalApplicationAsync();
            break;
        case 9:
            await ListServicePrincipalsWithExternalApplicationAsync();
            break;
        case 10:
            await ListApplicationsWithoutServicePrincipalAsync();
            break;
        case 11:
            // List ServicePrincipals
            await ListServicePrincipalsWithPublisherAsync();
            break;
        case 12:
            Console.WriteLine("Please enter an AppName to look for (emtpy for default):");
            string? displayName = Console.ReadLine();
            if (String.IsNullOrWhiteSpace(displayName))
            {
                displayName = settings.AppDisplayName ?? "Azure";
            }
            await ListSignInsAsync(displayName);
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}
void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task DisplayAccessTokenAsync()
{
    try
    {
        var userToken = await GraphHelper.GetUserTokenAsync();
        Console.WriteLine($"User token: {userToken}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user access token: {ex.Message}");
    }
}


async Task ListApplicationsAsync()
{
    try
    {
        var applications = await GraphHelper.ListApplicationsAsync();
        Console.WriteLine($"# Applications: {applications?.Value?.Count()}");
        if (null != applications?.Value)
        {
            foreach (var app in applications.Value)
            {
                StringBuilder sb = new StringBuilder("RequiredResourceAccess:");
                // List permission scopes an application MAY request
                // https://learn.microsoft.com/en-us/graph/api/resources/requiredresourceaccess
                if (null != app.RequiredResourceAccess)
                {
                    foreach (var rra in app.RequiredResourceAccess)
                    {
                        sb.Append(rra.ResourceAppId);
                        if (null != rra.ResourceAccess)
                        {
                            foreach (var ra in rra.ResourceAccess)
                            {
                                sb.Append($"{ra.Id}-{ra.Type}/");
                            }
                        }
                        sb.Append("-");
                    }
                }
                Console.WriteLine($"{app.DisplayName} | SignInAudience:{app.SignInAudience} | AppId:{app.AppId} | RequiredResourceAccess:{sb}");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting applications: {ex.Message}");
    }

}

async Task ListApplicationsWithoutServicePrincipalAsync()
{
    try
    {
        var applications = await GraphHelper.ListApplicationsWithoutServicePrincipalAsync();
        Console.WriteLine($"# Applications without ServicePrincipal: {applications?.Count()}");
        if (null != applications)
        {
            foreach (var app in applications)
            {
                Console.WriteLine($"{app.DisplayName} | SignInAudience:{app.SignInAudience} | AppId:{app.AppId}");
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting applications: {ex.Message}");
    }

}

async Task ListServicePrincipalsAsync(bool withoutMsApps = false)
{
    try
    {
        var servicePrincipals = withoutMsApps? await GraphHelper.ListServicePrincipalsWithoutMicrosoftAppsAsync() : await GraphHelper.ListServicePrincipalsAsync();
        Console.WriteLine($"# ServicePrincipals: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}
async Task ListServicePrincipalsWithPublisherAsync(bool withoutMsApps = false)
{
    try
    {
        var servicePrincipals = withoutMsApps? await GraphHelper.ListServicePrincipalsWithoutMicrosoftAppsAsync() : await GraphHelper.ListServicePrincipalsAsync();
        Console.WriteLine($"# ServicePrincipals: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                if (!String.IsNullOrEmpty(spn.VerifiedPublisher?.DisplayName))
                {
                    Console.WriteLine(await GetServicePrincipalAsString(spn));
                }
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}

async Task ListServicePrincipalsEnterpriseAsync()
{
    try
    {
        var servicePrincipals = await GraphHelper.ListServicePrincipalsEnterpriseAsync();
        Console.WriteLine($"# ServicePrincipals of type EntepriseApp: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}
async Task ListManagedIdentitiesAsync()
{
    try
    {
        var servicePrincipals = await GraphHelper.ListManagedIdentitiesAsync();
        Console.WriteLine($"# ManagedIdentities: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}
async Task ListServicePrincipalsWithInternalApplicationAsync()
{
    try
    {
        var servicePrincipals = await GraphHelper.ListServicePrincipalsWithInternalApplicationAsync();
        Console.WriteLine($"# ServicePrincipals with intenral application: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}

async Task ListServicePrincipalsWithExternalApplicationAsync()
{
    try
    {
        var servicePrincipals = await GraphHelper.ListServicePrincipalsWithExternalApplicationAsync();
        Console.WriteLine($"# ServicePrincipals with external application: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicwePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}

async Task ListServicePrincipalPermissionsAsync(string appName)
{
    try
    {
        var servicePrincipals = await GraphHelper.GetApplicatonPermissionsAsync(appName);
        Console.WriteLine($"# ServicePrincipals: {servicePrincipals?.Value?.Count()}");
        if (null != servicePrincipals?.Value)
        {
            foreach (var spn in servicePrincipals.Value)
            {
                Console.WriteLine(await GetServicePrincipalAsString(spn));
            }
        }
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"ODataError getting SignIns: {odataError.ToString()}");
        Console.WriteLine(odataError.Error?.Code);
        Console.WriteLine(odataError.Error?.Message);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting ServicePrincipals: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}

async Task<string> GetServicePrincipalAsString(ServicePrincipal spn)
{
    StringBuilder sb = new StringBuilder();
    string permissions = await GraphHelper.GetApplicatonPermissionsAsStringAsync(spn);
    string appRoles = await GraphHelper.GetAssignedUsersAsync(spn);
    string vendorType = String.Empty;
    string tags = String.Empty;
    if (null != spn.ServicePrincipalType)
    {
        if (spn.ServicePrincipalType == "ManagedIdentity")
        {
            vendorType += "ManagedIdentity";
        }
    }
    if (null != spn.Tags)
    {
        tags = String.Join('-', spn.Tags);
        if (spn.Tags.Contains("WindowsAzureActiveDirectoryIntegratedApp"))
        {
            vendorType += "EnterpriseApp-";
        }
    }
    if (null != spn.AppOwnerOrganizationId)
    {
        if (spn.AppOwnerOrganizationId == System.Guid.Parse(settings.TenantId ?? String.Empty))
        {
            vendorType +=  "Internal";
        }
        else
        {
            vendorType += "External";
        }
    }

    sb.Append($"{spn.DisplayName} | AppId:{spn.AppId} | Permissions:{permissions} | AppRoles:{appRoles} | AppOwnerTenant:{spn.AppOwnerOrganizationId} | Vendor: {vendorType} |  VerifiedPublisher:{spn.VerifiedPublisher?.DisplayName} | PrincipalType:{spn.ServicePrincipalType} | SignInAudience:{spn.SignInAudience} | Tags:{tags}" );
    return sb.ToString();
}

async Task ListSignInsAsync(string appName)
{
    try
    {
        var signInCollectionResponse = await GraphHelper.GetApplicatonSignInsAsync(appName);
        Console.WriteLine($"# SignIns: {signInCollectionResponse?.Value?.Count()}");
        if (null != signInCollectionResponse?.Value)
        {
            foreach (var signIn in signInCollectionResponse.Value)
            {
                Console.WriteLine($"{signIn.CreatedDateTime} | {signIn.UserDisplayName} | {signIn.UserPrincipalName} | {signIn.AppDisplayName} | {signIn.AppId} | {signIn.ClientAppUsed} ");
            }
        }
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"ODataError getting SignIns: {odataError.ToString()}");
        Console.WriteLine(odataError.Error?.Code);
        Console.WriteLine(odataError.Error?.Message);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting SignIns: {ex.Message}");
        if (null != ex.InnerException)
        {
            Console.WriteLine(ex.InnerException.Message);
        }
    }
}

