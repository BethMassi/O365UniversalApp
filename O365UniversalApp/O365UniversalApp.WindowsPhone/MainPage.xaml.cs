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
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace O365UniversalApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page, IWebAuthenticationContinuable
    {
        AuthenticationContext authContext = null;
        List<ServiceCapabilityInfo> appCapabilities = new List<ServiceCapabilityInfo>();

        public MainPage()
        {
            this.InitializeComponent();

            this.NavigationCacheMode = NavigationCacheMode.Required;
        }

        /// <summary>
        /// Invoked when this page is about to be displayed in a Frame.
        /// </summary>
        /// <param name="e">Event data that describes how this page was reached.
        /// This parameter is typically used to configure the page.</param>
        protected async override void OnNavigatedTo(NavigationEventArgs e)
        {
            // TODO: Prepare page for display here.

            // TODO: If your application contains multiple pages, ensure that you are
            // handling the hardware Back button by registering for the
            // Windows.Phone.UI.Input.HardwareButtons.BackPressed event.
            // If you are using the NavigationHelper provided by some templates,
            // this event is handled for you.
            authContext = await AuthenticationContext.CreateAsync(AuthenticationHelper.Authority);
        }

        public async void ContinueWebAuthentication(Windows.ApplicationModel.Activation.WebAuthenticationBrokerContinuationEventArgs args)
        {
            await authContext.ContinueAcquireTokenAsync(args);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            authContext.AcquireTokenAndContinue(AuthenticationHelper.DiscoveryServiceResourceId,
                AuthenticationHelper.ClientId,
                AuthenticationHelper.AppRedirectURI,
                async (AuthenticationResult result) =>
                {
                    AuthenticationHelper.LastTenantId = result.TenantId;

                    if (appCapabilities.Count == 0)
                    {
                        appCapabilities = await Office365APIHelper.DiscoverAppCapabilities(result.AccessToken);

                        foreach (var capabilityInfo in appCapabilities)
                        {
                            MessageDialog mg = new MessageDialog(String.Format("{0} -> {1}, {2}", capabilityInfo.Name, capabilityInfo.ServiceEndpointUri, capabilityInfo.ServiceResourceId));
                            await mg.ShowAsync();
                        }
                    }

                    await GetContacts();
                });
        }

        private async Task GetContacts()
        {
            var contactsCapabilityInfo = appCapabilities.FirstOrDefault(s => s.Name == "Contacts");

            if (contactsCapabilityInfo != null)
            {
                var authResult = await authContext.AcquireTokenSilentAsync(contactsCapabilityInfo.ServiceResourceId, AuthenticationHelper.ClientId);

                var myContacts = await Office365APIHelper.GetContacts(contactsCapabilityInfo, authResult.AccessToken);

                foreach (var myContact in myContacts)
                {
                    MessageDialog mg = new MessageDialog(String.Format("{0} - {1}", myContact.Name, myContact.JobTitle));
                    await mg.ShowAsync();
                }

            }
        }
    }
}
