// See https://aka.ms/new-console-template for more information
// Created via tutorial https://learn.microsoft.com/en-us/graph/tutorials/dotnet?tabs=aad
Console.WriteLine(".NET Graph AppRegistrations\n");

var settings = Settings.LoadSettings();

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
                Console.WriteLine($"{app.DisplayName} - {app.SignInAudience}");
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
                string permissions = String.Empty;
                if ( null != spn.Oauth2PermissionScopes)
                {
                    foreach (var scope in spn.Oauth2PermissionScopes)
                    {
                        permissions += scope.Type + ":" + scope.Value;
                    }
                }
                Console.WriteLine($"{spn.DisplayName} - {spn.AppId} - Permissions {permissions} - App Owner Tenant {spn.AppOwnerOrganizationId} - {spn.ServicePrincipalType}" );
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