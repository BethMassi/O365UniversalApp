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
using System.Linq;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace O365UniversalApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
        public sealed partial class MainPage : Page
        {
            public MainPage()
            {
                this.InitializeComponent();
            }

            private async void Button_Click(object sender, RoutedEventArgs e)
            {
                var authResult = await AuthenticationHelper.GetAccessToken(AuthenticationHelper.DiscoveryServiceResourceId);

                var capabilities = await Office365APIHelper.DiscoverAppCapabilities(authResult.AccessToken);
                
                foreach (var capabilityInfo in capabilities)
                {
                    MessageDialog mg = new MessageDialog(String.Format("{0} -> {1}, {2}",capabilityInfo.Name, capabilityInfo.ServiceEndpointUri,capabilityInfo.ServiceResourceId));
                    await mg.ShowAsync();
                }

                var contactsCapabilityInfo = capabilities.FirstOrDefault(s => s.Name == "Contacts");
                if (contactsCapabilityInfo != null)
                {
                    authResult = await AuthenticationHelper.GetAccessToken(contactsCapabilityInfo.ServiceResourceId);
                    if (authResult.Status == AuthenticationStatus.Success)
                    {
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
}
