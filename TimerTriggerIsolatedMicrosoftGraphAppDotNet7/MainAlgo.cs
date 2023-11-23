using Azure.Storage.Blobs;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using System.ComponentModel;
using System.Reflection.Metadata;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using static System.Reflection.Metadata.BlobBuilder;

namespace TimerTriggerIsolatedMicrosoftGraphAppDotNet7
{
    internal class MainAlgo
    {
        public async Task ProgramRunAsync(string storageConnectionString, CloudStorageAccount cloudStorageAccount, CloudBlobClient cloudBlobClient, CloudBlobContainer cloudBlobContainer, CloudBlockBlob cloudBlockBlob, ILogger log)
        {
            log.LogWarning(".NET Graph App-only Tutorial\n");

            // Initialize Graph
            Settings settings = Settings.LoadSettings();
            InitializeGraph(settings);

            log.LogInformation("5. Test the custom Group Graph call");
            await ListMemberGraphCallAsync(storageConnectionString, cloudStorageAccount, cloudBlobClient, cloudBlobContainer, cloudBlockBlob, log);
        }

        void InitializeGraph(Settings settings)
        {
            GraphHelper.InitializeGraphForAppOnlyAuth(settings);
        }

        async Task DisplayAccessTokenAsync(ILogger log)
        {
            try
            {
                var appOnlyToken = await GraphHelper.GetAppOnlyTokenAsync();
                log.LogInformation($"App-only token: {appOnlyToken}");
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error getting app-only access token: {ex.Message}");
            }
        }

        static async Task ListMemberGraphCallAsync(string storageConnectionString, CloudStorageAccount cloudStorageAccount, CloudBlobClient cloudBlobClient, CloudBlobContainer cloudBlobContainer, CloudBlockBlob cloudBlockBlob, ILogger log)
        {
            try
            {
                StringBuilder json = new();
                var groupIds = await GraphHelper.GetListGroupsIdsAsync();

                if (groupIds?.Value == null)
                {
                    log.LogInformation("No results returned.");
                    return;
                }

                json.Append('[');
                // Output each users's details
                foreach (var groupId in groupIds.Value)
                {
                    string? id = groupId.Id.ToString();

                    var groupProperties = await GraphHelper.GetGroupsPropertiesAsync(id);

                    if (groupProperties == null)
                    {
                        log.LogInformation("No results returned.");
                        return;
                    }
                    json.Append("{");
                    json.Append($" \"id\": \"{groupProperties.Id ?? "null"}\"" + ",");
                    json.Append($" \"displayName\": \"{groupProperties.DisplayName ?? "null"}\"" + ",");
                    json.Append($" \"createdDateTime\": \"{groupProperties.CreatedDateTime.ToString() ?? "null"}\"" + ",");
                    json.Append($" \"owner\": \"{groupProperties.Mail ?? "null"}\"" + ",");

                    var owners = await GraphHelper.GetGroupsOwnersAsync(id);
                    if (owners == null)
                    {
                        log.LogInformation("No results returned.");
                        return;
                    }
                    json.Append(" \"owners\":[");

                    int indexOwners = 0;
                    int countOwners = owners.Value.Count;

                    // Output each group member's details
                    foreach (dynamic owner in owners.Value)
                    {
                        json.Append("{");
                        json.Append($" \"memberId\": \"{owner.Id ?? "null"}\"" + ",");
                        json.Append($" \"memberName\": \"{owner.DisplayName ?? "null"}\"" + ",");
                        json.Append($" \"memberMail\": \"{owner.Mail ?? "null"}\"");
                        json.Append("}");

                        if (indexOwners < countOwners - 1)
                        {
                            json.Append(',');
                        }

                        indexOwners++;
                    }
                    json.Append("],");

                    var groupMembers = await GraphHelper.GetListGroupsMembersAsync(id);

                    if (groupMembers?.Value == null)
                    {
                        log.LogInformation("No results returned.");
                        return;
                    }

                    json.Append(" \"members\":[");
                    int indexGroupMembers = 0;
                    int countGroupMembers = groupMembers.Value.Count;

                    // Output each group member's details
                    foreach (var groupMember in groupMembers.Value)
                    {
                        json.Append("{");
                        json.Append($" \"memberId\": \"{groupMember.Id ?? "null"}\"" + ",");
                        json.Append($" \"memberName\": \"{groupMember.DisplayName ?? "null"}\"" + ",");
                        json.Append($" \"memberMail\": \"{groupMember.Mail ?? "null"}\"");
                        json.Append("}");

                        if (indexGroupMembers < countGroupMembers - 1)
                        {
                            json.Append(',');
                        }

                        indexGroupMembers++;
                    }
                    json.Append("],");

                    json.Append(" \"groupTypes\":[");
                    foreach (var groupProperty in groupProperties.GroupTypes)
                    {
                        json.Append($" \"{groupProperty}\" ");

                    }
                    json.Append("],");
                    json.Append($" \"expirationDateTime\":{(groupProperties.ExpirationDateTime.HasValue ? groupProperties.ExpirationDateTime.Value.ToString("o") : "null")}");
                    json.Append("},");
                }
                int lastCommaPositionOuter = json.ToString().LastIndexOf(",");
                json.Remove(lastCommaPositionOuter, 1);
                json.Append(']');
                log.LogInformation(json.ToString());
                try
                {
                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true
                    };

                    // Convert StringBuilder to TextWriter
                    var document = JsonDocument.Parse(json.ToString());

                    // Serialize the JSON with indented formatting
                    string formattedJson = System.Text.Json.JsonSerializer.Serialize(document.RootElement, options);

                    // Overwrite the existing JSON file with the updated content
                    await cloudBlockBlob.UploadTextAsync(formattedJson);
                    
                    log.LogWarning("JSON file updated and uploaded successfully.");
                }
                catch (Exception ex)
                {
                    log.LogError(ex, "An error occurred while updating and uploading the JSON file.");
                }

            }
            catch (Exception ex)
            {
                log.LogInformation($"Error making Graph Call: {ex.Message}");
            }
        }
    }
}

