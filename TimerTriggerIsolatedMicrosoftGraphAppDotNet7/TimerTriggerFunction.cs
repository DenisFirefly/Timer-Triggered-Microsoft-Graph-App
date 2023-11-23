using Azure.Storage.Blobs.Specialized;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage;

namespace TimerTriggerIsolatedMicrosoftGraphAppDotNet7
{
    public class TimerTriggerFunction
    {
        private readonly ILogger _logger;

        public TimerTriggerFunction(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<TimerTriggerFunction>();
        }

        [Function("Function1")]
        public async Task RunAsync([TimerTrigger("0 * 9 * Jun Wed", RunOnStartup = true)] MyInfo myTimer)
        {
            _logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            _logger.LogInformation($"Next timer schedule at: {myTimer.ScheduleStatus.Next}");

            // Retrieve the connection string for the Azure Storage account
            string? storageConnectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage", EnvironmentVariableTarget.Process);
            _logger.LogInformation($"Key: {storageConnectionString}");

            // Create a CloudStorageAccount object
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageConnectionString);

            // Create a CloudBlobClient object
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            // Create a reference to the container where the JSON file will be uploaded
            CloudBlobContainer container = blobClient.GetContainerReference("mycontainer");

            // Create a reference to the JSON file in the container
            CloudBlockBlob blob = container.GetBlockBlobReference("PT1H.json");

            MainAlgo program = new();
            await program.ProgramRunAsync(storageConnectionString, storageAccount, blobClient, container, blob,_logger);
        }

    }

    public class MyInfo
    {
        public required MyScheduleStatus ScheduleStatus { get; set; }

        public bool IsPastDue { get; set; }
    }

    public class MyScheduleStatus
    {
        public DateTime Last { get; set; }

        public DateTime Next { get; set; }

        public DateTime LastUpdated { get; set; }
    }
}
