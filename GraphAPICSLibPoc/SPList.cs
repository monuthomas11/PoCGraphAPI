using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

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
            var result = await _graphServiceClient.Drives["{drive-id}"].Items["{driveItem-id}"].Children.GetAsync();
            //var drive = await _graphServiceClient.Me.Drive.GetAsync();
            return JsonSerializer.Serialize(result);
        }
    }
    public class ItemModel
    {
        public string Title { get; set; }
        public string Name { get; set; }
    }
}
