using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GraphAPICSLibPoc
{
    internal class ClientCredentialProvider : IAuthenticationProvider
    {
        private readonly IConfidentialClientApplication _confidentialClientApplication;
        public ClientCredentialProvider(IConfidentialClientApplication confidentialClientApplication)
        {
            _confidentialClientApplication = confidentialClientApplication;
        }

        public async Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object> additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            var authResult = await _confidentialClientApplication
                            .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            request.Headers.Add("Authorization", $"Bearer {authResult.AccessToken}");
        }
    }
}
