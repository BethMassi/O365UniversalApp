//----------------------------------------------------------------------------------------------
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.UI.Popups;

namespace O365UniversalApp
{
    public static partial class AuthenticationHelper
    {
        static AuthenticationContext authContext = null;

        public static async Task<AuthenticationResult> GetAccessToken(string serviceResourceId)
        {
            if (authContext == null)
            {
                var authority = String.Format(AuthorityFormat, "common");
                authContext = new AuthenticationContext(authority);
            }

            if (!String.IsNullOrEmpty(LastTenantId))
            {
                var authority = String.Format(AuthorityFormat, LastTenantId);
                authContext = new AuthenticationContext(authority);
            }

            //authContext.UseCorporateNetwork = true;
           
            var redirectUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();

            var authResult = await authContext.AcquireTokenAsync(serviceResourceId, ClientId, redirectUri);

            LastTenantId = authResult.TenantId;

            if (authResult.Status != AuthenticationStatus.Success)
            {
                if (authResult.Error == "authentication_canceled")
                {
                    // The user cancelled the sign-in, no need to display a message.
                }
                else
                {
                    MessageDialog dialog = new MessageDialog(string.Format("If the error continues, please contact your administrator.\n\nError: {0}\n\n Error Description:\n\n{1}", authResult.Error, authResult.ErrorDescription), "Sorry, an error occurred while signing you in.");
                    await dialog.ShowAsync();
                }
            }

            return authResult;
        }      
    }
}
