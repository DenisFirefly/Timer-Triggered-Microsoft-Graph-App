using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text.Json;

namespace TimerTriggerIsolatedMicrosoftGraphAppDotNet7;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // App-ony auth token credential
    private static ClientSecretCredential? _clientSecretCredential;
    // Client configured with app-only authentication
    private static GraphServiceClient? _appClient;

    public static void InitializeGraphForAppOnlyAuth(Settings settings)
    {
        _settings = settings;

        // Ensure settings isn't null
        _ = settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        _settings = settings;

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] { "https://graph.microsoft.com/.default" });
        }
    }

    public static async Task<string> GetAppOnlyTokenAsync()
    {
        // Ensure credential isn't null
        _ = _clientSecretCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        // Request token with given scopes
        var context = new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
        var response = await _clientSecretCredential.GetTokenAsync(context);
        return response.Token;
    }

    public static Task<UserCollectionResponse?> GetUsersAsync()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        #region OLD
        var result = _appClient.Users.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "id", "mail" };
            // Get at most 25 results
            config.QueryParameters.Top = 25;
            // Sort by display name
            config.QueryParameters.Orderby = new[] { "displayName" };
        });
        // Serialize the result to JSON
        var options = new JsonSerializerOptions
        {
            WriteIndented = true // to format the JSON with indentation
        };
        string json = JsonSerializer.Serialize(result, options);
        Console.WriteLine(json);
        #endregion

        return _appClient.Users.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "id", "mail" };
            // Get at most 25 results
            config.QueryParameters.Top = 25;
            // Sort by display name
            config.QueryParameters.Orderby = new[] { "displayName" };
        });
    }

    // This function serves as a playground for testing Graph snippets
    // or other code
    public static Task<GroupCollectionResponse?> MakeGraphCallAsync()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        #region OLD
        /*var result = _appClient.Groups.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "Id" };
            // Get at most 25 results
            config.QueryParameters.Top = 25;
            // Sort by display name
            config.QueryParameters.Orderby = new[] { "DisplayName" };
            config.Headers.Add("ConsistencyLevel", "eventual");
            config.Headers.Add("Content-Type", "application/json");
        });
        // Serialize the result to JSON
        var options = new JsonSerializerOptions
        {
            WriteIndented = true // to format the JSON with indentation
        };
        string json = JsonSerializer.Serialize(result, options);
        Console.WriteLine(json);*/
        #endregion

        return _appClient.Groups.GetAsync((config) =>
        {
            // Only request specific properties
            config.QueryParameters.Select = new[] { "displayName", "id", "mail" };
            // Get at most 25 results
            config.QueryParameters.Top = 25;
            // Sort by display name
            config.QueryParameters.Orderby = new[] { "displayName" };
        });
    }

    public static Task<GroupCollectionResponse?> GetListGroupsIdsAsync()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Groups.GetAsync();

    }

    public static Task<Microsoft.Graph.Models.Group?> GetGroupsPropertiesAsync(string id)
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Groups[id].GetAsync((config) =>
        {
            config.QueryParameters.Select = new string[] { "id", "displayName", "description", "createdDateTime", "mail", "groupTypes", "expirationDateTime" };
        });
    }

    public static Task<DirectoryObjectCollectionResponse?> GetGroupsOwnersAsync(string id)
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Groups[id].Owners.GetAsync((config) =>
        {
            config.QueryParameters.Select = new string[] { "id", "displayName", "description", "createdDateTime", "mail", "groupTypes", "expirationDateTime" };
        });


    }

    public static Task<UserCollectionResponse?> GetListGroupsMembersAsync(string id)
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Groups[id].Members.GraphUser.GetAsync((config) =>
        {
            config.QueryParameters.Select = new string[] { "id", "displayName", "mail" };
            config.Headers.Add("ConsistencyLevel", "eventual");
        });

    }

    public static Task<UserCollectionResponse?> GetListUsersAsync()
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Users.GetAsync((config) =>
        {
            config.QueryParameters.Select = new string[] { "id", "displayName", "jobTitle", "mail", "BusinessPhones", "mobilePhone" };
            config.Headers.Add("ConsistencyLevel", "eventual");
        });

    }

    public static Task<DirectoryObjectCollectionResponse?> GetListUsersMemberOfAsync(string id)
    {
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Groups[id].MemberOf.GetAsync((config) =>
        {
            config.QueryParameters.Select = new string[] { "id", "createdDateTime", "creationOptions", "description", "displayName", "groupTypes" };
            config.Headers.Add("ConsistencyLevel", "eventual");
        });

    }
}
