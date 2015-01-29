using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TRO365APIDemo
{
    internal class AuthHelper
    {
        AuthenticationResult authResult;
        System.Threading.ManualResetEventSlim resultReady;

        public AuthHelper()
        {
            resultReady = new System.Threading.ManualResetEventSlim(false);
        }

        private void GetAccessToken(AuthenticationResult result)
        {
            authResult = result;
            resultReady.Set();
        }

        public async Task<AuthenticationResult> GetAccessTokenForService(AuthenticationContext authContext,string resource)
        {
            AuthenticationResult result = await authContext.AcquireTokenSilentAsync(AuthenticationHelper.DiscoveryServiceResourceId, AuthenticationHelper.ClientId);

            if (result != null && result.Status == AuthenticationStatus.Success)
            {
                return result;
            }
            else
            {
                authContext.AcquireTokenAndContinue(resource, AuthenticationHelper.ClientId, new Uri(""), GetAccessToken);
                while (!resultReady.Wait(new TimeSpan(0,0,0,0,100)))
                {
                    await Task.Delay()
                }
                
                return authResult;
            }
        }
    }
}
