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

using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace O365UniversalApp
{
    public static class Office365APIHelper
    {      
        public static async Task<List<ServiceCapabilityInfo>> DiscoverAppCapabilities(string accesstoken)
        {
            List<ServiceCapabilityInfo> serviceCapabilitiesInfo = new List<ServiceCapabilityInfo>();

            DiscoveryClient discoveryClient = new DiscoveryClient
            (
                () =>
                {
                    return accesstoken;
                }
            );

            var capabilities = await discoveryClient.DiscoverCapabilitiesAsync();

            
            foreach (var capability in capabilities)
            {
                serviceCapabilitiesInfo.Add(new ServiceCapabilityInfo
                {
                    Name = capability.Key,
                    ServiceEndpointUri = capability.Value.ServiceEndpointUri,
                    ServiceResourceId = capability.Value.ServiceResourceId
                });
            }


            return serviceCapabilitiesInfo;
        }

        public static async Task<List<MyContact>> GetContacts(ServiceCapabilityInfo contactsServiceCapability, string accesstoken)
        {
            List<MyContact> myContacts = new List<MyContact>();

            OutlookServicesClient outlookClient = new OutlookServicesClient(contactsServiceCapability.ServiceEndpointUri,
                async () =>
                {
                    return await Task.FromResult(accesstoken);
                });

            var contactsResult = await outlookClient.Me.Contacts.ExecuteAsync();
            var contacts = contactsResult.CurrentPage;
            foreach (var contact in contacts)
            {
                myContacts.Add(new MyContact
                {
                    Name = contact.DisplayName,
                    JobTitle = contact.JobTitle
                });
            }

            return myContacts;
        }       
        
    }
}
