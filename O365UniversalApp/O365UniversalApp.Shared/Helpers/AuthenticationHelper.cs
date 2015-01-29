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

using Refractored.Xam.Settings;
using Refractored.Xam.Settings.Abstractions;
using System;

namespace O365UniversalApp
{
    public static partial class AuthenticationHelper
    {        
        public const string AuthorityFormat = "https://login.windows.net/{0}/";
        
        public static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");

        public static readonly string DiscoveryServiceResourceId = "https://api.office.com/discovery/";

        private static string _authority = String.Empty;        

        private const string LastTenantIdKey = "last_authority_key";
        private static readonly string LastTenantIdDefault = "common";

        private static ISettings AppSettings
        {
            get
            {
                return CrossSettings.Current;
            }
        }

        public static string LastTenantId
        {
            get
            {
                return AppSettings.GetValueOrDefault(LastTenantIdKey, LastTenantIdDefault);
            }
            set
            {
                AppSettings.AddOrUpdateValue(LastTenantIdKey, value);
            }
        }

       
#if NETFX_CORE

        public static readonly string ClientId = App.Current.Resources["ida:ClientID"].ToString();

        public static Uri AppRedirectURI
        {
            get
            {
                return Windows.Security.Authentication.Web.WebAuthenticationBroker.GetCurrentApplicationCallbackUri();
            }
        }
#endif
     
        public static string Authority
        {    
            get
            {
                if (!String.IsNullOrEmpty(LastTenantId))
                {
                    _authority = String.Format(AuthorityFormat, LastTenantId);
                }
                else
                {
                    _authority = String.Format(AuthorityFormat, "common");
                }

                return _authority;
            }
        }

        //private static ApplicationDataContainer _tenantIdSettings = ApplicationData.Current.LocalSettings;
        //public static string TenantId
        //{
        //    get
        //    {
        //        if (_tenantIdSettings.Values.ContainsKey("LastAuthority") && _tenantIdSettings.Values["LastAuthority"] != null)
        //        {
        //            return _tenantIdSettings.Values["LastAuthority"].ToString();
        //        }
        //        else
        //        {
        //            return string.Empty;
        //        }

        //    }

        //    set
        //    {
        //        _tenantIdSettings.Values["LastAuthority"] = value;
        //    }
        //}

        
    }
}
