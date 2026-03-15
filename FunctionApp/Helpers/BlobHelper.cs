using System;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;

namespace SlideTemplateFiller.Functions.Helpers
{
    public class BlobHelper
    {
        private readonly BlobServiceClient _serviceClient;
        private readonly string _accountName;
        private readonly string _accountKey;

        public BlobHelper(string connectionString)
        {
            _serviceClient = new BlobServiceClient(connectionString);

            // parse account name/key from connection string to sign SAS tokens
            // connection string looks like: "DefaultEndpointsProtocol=https;AccountName=xxx;AccountKey=yyy;EndpointSuffix=core.windows.net"
            var parts = connectionString.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var p in parts)
            {
                var kv = p.Split('=', 2);
                if (kv.Length != 2) continue;
                var key = kv[0].Trim();
                var val = kv[1].Trim();
                if (string.Equals(key, "AccountName", StringComparison.OrdinalIgnoreCase)) _accountName = val;
                if (string.Equals(key, "AccountKey", StringComparison.OrdinalIgnoreCase)) _accountKey = val;
            }
        }

        public async Task DownloadBlobToFileAsync(string containerName, string blobName, string localPath)
        {
            var container = _serviceClient.GetBlobContainerClient(containerName);
            await container.CreateIfNotExistsAsync();
            var blob = container.GetBlobClient(blobName);
            await blob.DownloadToAsync(localPath);
        }

        public async Task UploadFileAsync(string containerName, string blobName, string filePath)
        {
            var container = _serviceClient.GetBlobContainerClient(containerName);
            await container.CreateIfNotExistsAsync();
            var blob = container.GetBlobClient(blobName);
            using var fs = File.OpenRead(filePath);
            await blob.UploadAsync(fs, overwrite: true);
        }

        public Uri GenerateReadSasUri(string containerName, string blobName, int minutes = 30)
        {
            var container = _serviceClient.GetBlobContainerClient(containerName);
            var blob = container.GetBlobClient(blobName);

            // Build sas
            var sasBuilder = new BlobSasBuilder
            {
                BlobContainerName = containerName,
                BlobName = blobName,
                Resource = "b",
                ExpiresOn = DateTimeOffset.UtcNow.AddMinutes(minutes)
            };
            sasBuilder.SetPermissions(BlobSasPermissions.Read);

            // Use account key to sign SAS
            if (string.IsNullOrEmpty(_accountName) || string.IsNullOrEmpty(_accountKey))
            {
                // If no key available, attempt client-side generation if the client supports it
                if (blob.CanGenerateSasUri)
                {
                    return blob.GenerateSasUri(sasBuilder);
                }
                throw new InvalidOperationException("Storage account key not found in connection string; cannot generate SAS token.");
            }

            var credential = new StorageSharedKeyCredential(_accountName, _accountKey);
            var sasQuery = sasBuilder.ToSasQueryParameters(credential).ToString();
            var uriBuilder = new UriBuilder(blob.Uri)
            {
                Query = sasQuery
            };
            return uriBuilder.Uri;
        }
    }
}