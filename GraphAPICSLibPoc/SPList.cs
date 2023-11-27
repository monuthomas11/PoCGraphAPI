using Microsoft.Graph;
using Microsoft.Graph.DeviceManagement.UserExperienceAnalyticsWorkFromAnywhereMetrics.Item.MetricDevices.Item;
using Microsoft.Graph.Drives.Item.Items.Item.Restore;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Graph.Models.Security;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;

namespace GraphAPICSLibPoc
{
    public class SPList
    {
        private IConfidentialClientApplication _confidentialClientApplication;
        private ClientCredentialProvider _authProvider;
        private GraphServiceClient _graphServiceClient;

        public SPList(string clientId, string tenantId, string clientSecret)
        {
            _confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();
            _authProvider = new ClientCredentialProvider(_confidentialClientApplication);
            _graphServiceClient = new GraphServiceClient(_authProvider);
        }
        public async Task<ItemModel> GetItem(string siteId, string listId, string itemId)
        {
            ItemModel model = null;

            var item = await _graphServiceClient.Sites[siteId].Lists[listId].Items[itemId].GetAsync();
            model = new ItemModel { Name = item.Fields.AdditionalData["Name"].ToString(), Title = item.Fields.AdditionalData["Title"].ToString() };

            return model;
        }
        public async Task<string> GetMyDrive()
        {
            var me = await _graphServiceClient.Groups["{12e7dae8-2c16-4ed3-8015-cfe7178b5bd5}"].Drive.GetAsync(); 
            //var drive = await _graphServiceClient.Me.Drive.GetAsync();
            return JsonSerializer.Serialize(me);
        }

        public async Task<string> GetDriveItems()
        {
            var result = await _graphServiceClient
                //.Groups["12e7dae8-2c16-4ed3-8015-cfe7178b5bd5"]
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFMBDYD5MEMOIZC3T437XYQGD2KN"]
                .Children
                .GetAsync();
            //var result = await _graphServiceClient.Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"].Items["{012FCIDFL2ANWARFF4GZFKQVGYRJGC5X7Q}"].Children.GetAsync();
           // var result = await _graphServiceClient.Drives["{drive-id}"].Items["{driveItem-id}"].Children.GetAsync();
            return JsonSerializer.Serialize(result);
        }

        public async Task<string> UpdateFileName()
        {
            var requestBody = new DriveItem
            {
                Name = "FileUploaded.png",
            };

            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            var result = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFPWLLKFUATKBJF2QMPST4MFPSUU"] // update folder nams/properties as well as file names/properties
                .PatchAsync(requestBody);

            return JsonSerializer.Serialize(result);
        }

        public async Task<string> UploadFile()
        {
            string result = "";
            using (var fileStream = File.OpenRead($"C:\\Users\\monut\\OneDrive\\Pictures\\Screenshots\\2017-07-14.png"))
            {

                // Use properties to specify the conflict behavior
                //using (var DriveUpload = Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession)
                var uploadSessionRequestBody = new DriveUpload.CreateUploadSessionPostRequestBody
                {
                    Item = new DriveItemUploadableProperties
                    {
                        AdditionalData = new Dictionary<string, object>
                        {
                            { "@microsoft.graph.conflictBehavior", "replace" },
                        },
                    },
                };

                // Create the upload session
                // itemPath does not need to be a path to an existing item
                //var myDrive = await _graphServiceClient.Me.Drive.GetAsync();
                var myDrive = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"].GetAsync();

                var uploadSession = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                    .Items["012FCIDFMBDYD5MEMOIZC3T437XYQGD2KN"]
                    .ItemWithPath("/drives/b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-/root:/DemoFolder/FileUploaded.png") //provide filename for the new file 
                    .CreateUploadSession
                    .PostAsync(uploadSessionRequestBody);

                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask = new LargeFileUploadTask<DriveItem>(
                    uploadSession, fileStream, maxSliceSize, _graphServiceClient.RequestAdapter);

                var totalLength = fileStream.Length;
                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"Uploaded {prog} bytes of {totalLength} bytes");
                });

                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                    result = JsonSerializer.Serialize(uploadResult);
                    Console.WriteLine(uploadResult.UploadSucceeded ?
                        $"Upload complete, item ID: {uploadResult.ItemResponse.Id}" :
                        "Upload failed");
                }
                catch (ODataError ex)
                {
                    Console.WriteLine($"Error uploading: {ex.Error?.Message}");
                }
            }
            return result;

        }

        public async Task<string> DeleteFile() 
        {
            string result = "";
            var files = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFMBDYD5MEMOIZC3T437XYQGD2KN"]
                .Children
                .GetAsync();
            foreach ( var file in files.Value )
            {
                if (file.Name == "sample.pdf")
                {
                    var fileId = file.Id;
                    await _graphServiceClient
                                    .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                                    .Items[fileId] // update folder nams/properties as well as file names/properties
                                    .DeleteAsync();
                    result = "DELETED ID: " + fileId;
                    break;
                    
                }
            }
            return result;
        }

        //not working - not able to retrieve id for item in recycle bin
        public async Task<string> RestoreDeletedFile()
        {
            // Code snippets are only available for the latest version. Current version is 5.x

            // Dependencies
            //using Microsoft.Graph.Drives.Item.Items.Item.Restore;
            //using Microsoft.Graph.Models;

            var requestBody = new RestorePostRequestBody
            {
                ParentReference = new ItemReference
                {
                    Id = "012FCIDFMBDYD5MEMOIZC3T437XYQGD2KN",
                },
                Name = "SampleRestored.pdf",
            };

            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            var result = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFMNOPXB64CZDNGKTRWQSJ7QLSPB"]
                //.Restore.PostAsync(requestBody);
                .GetAsync();

            return JsonSerializer.Serialize(result);


        }

        public async Task<string> MoveItem()
        {
            var driveFolders = await _graphServiceClient.Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFN6Y2GOVW7725BZO354PWSELRRZ"].Children.GetAsync();

            //var items = await _graphServiceClient
            //    //.Sites["48c3dbd7-fe08-418f-a633-8c75e00b277b"]
            //    .Groups["12e7dae8-2c16-4ed3-8015-cfe7178b5bd5"]
            //    //.Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
            //        .Drive.
            string newFolderId = driveFolders.Value.Where(x=>x.Name == "TestFolder").Select(x=>x.Id).FirstOrDefault();

            var files = await _graphServiceClient
                .Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items["012FCIDFMBDYD5MEMOIZC3T437XYQGD2KN"]
                .Children
                .GetAsync();
                //.GetAsync((requestConfiguration) =>
                //{
                //    requestConfiguration.QueryParameters.Search = "\"NameUpdated\"";
                //}); // not filtering on search parameter

            string fileId = files.Value.Where(x => x.Name == "sample.pdf").Select(x => x.Id).FirstOrDefault();

            var requestBody = new DriveItem
            {
                ParentReference = new ItemReference
                {
                    Id = newFolderId
                },
                Name = "movedsample.pdf"
            };

            // To initialize your graphClient, see https://learn.microsoft.com/en-us/graph/sdks/create-client?from=snippets&tabs=csharp
            var result = await _graphServiceClient.Drives["b!19vDSAj-j0GmM4x14Asne4U9MMkh_SxItknxyVfMN15Bh7Yy1yRoRryIS3mrrSo-"]
                .Items[fileId].PatchAsync(requestBody);

            return JsonSerializer.Serialize(result);
        }

        //public async Task<DriveItemC>
    }
    public class ItemModel
    {
        public string Title { get; set; }
        public string Name { get; set; }
    }
}
