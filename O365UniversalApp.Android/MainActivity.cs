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


using Android.App;
using Android.Content;
using Android.OS;
using Android.Runtime;
using Android.Widget;
using Com.Microsoft.Aad.Adal;
using System;
using System.Collections.Generic;
using System.Linq;

namespace O365UniversalApp.Android
{
    [Activity(Label = "TRO365APIDemo.Android", MainLauncher = true, Icon = "@drawable/icon")]
    public class MainActivity : Activity
    {
        AuthenticationContext authContext;
        DefaultTokenCacheStore LocalAccountCache;
        internal UserInfo AadUserInfo;

        const string clientId = "6c1cdb44-d24e-41ba-ab22-257055ac955a";
        const string appRedirectUri = "https://xamarin-android-app";
            
        protected override void OnCreate(Bundle bundle)
        {
            base.OnCreate(bundle);

            // Set our view from the "main" layout resource
            SetContentView(Resource.Layout.Main);

            // Get our button from the layout resource,
            // and attach an event to it
            Button button = FindViewById<Button>(Resource.Id.MyButton);

            DefaultTokenCacheStore authTokenCache = new DefaultTokenCacheStore(this);
            authContext = new AuthenticationContext(this, AuthenticationHelper.Authority, false, authTokenCache);

            button.Click += button_Click;
        }

        void button_Click(object sender, EventArgs e)
        {
            AuthenticationResult authenticationResult = null;
            try
            {
                authenticationResult = authContext.AcquireTokenSilentSync(AuthenticationHelper.DiscoveryServiceResourceId, clientId, AadUserInfo.UserId);

            }
            catch (Exception exception)
            {
                //needs prompt
            }

            if (authenticationResult == null || authenticationResult.Status != AuthenticationResult.AuthenticationStatus.Succeeded)
            {
                authContext.AcquireToken(this, AuthenticationHelper.DiscoveryServiceResourceId, clientId, appRedirectUri, PromptBehavior.Auto, new TestCallback(this));
            }
            else
            {
                GetContacts();
            }
        }

        async void GetContacts()
        {
            var authResultDiscovery = authContext.AcquireTokenSilentSync(AuthenticationHelper.DiscoveryServiceResourceId, clientId, AadUserInfo.UserId);
            List<ServiceCapabilityInfo> appCapabilities = await Office365APIHelper.DiscoverAppCapabilities(authResultDiscovery.AccessToken);

            ServiceCapabilityInfo contactsCapabilitiyInfo = appCapabilities.FirstOrDefault(s => s.Name == "Contacts");
            var authResultOutlook = authContext.AcquireTokenSilentSync(contactsCapabilitiyInfo.ServiceResourceId, clientId, AadUserInfo.UserId);

            var myContacts = await Office365APIHelper.GetContacts(contactsCapabilitiyInfo, authResultOutlook.AccessToken);
            foreach (var myContact in myContacts)
            {

                AlertDialog.Builder builder = new AlertDialog.Builder(this);
                builder.SetMessage(myContact.Name);
                builder.SetTitle("Contact");
                builder.SetCancelable(false);
                builder.SetPositiveButton("OK", (sender, args) => { });
                builder.Create().Show();
            }

            if (AadUserInfo == null)
            {
                AadUserInfo = new UserInfo();
            }
        }

        protected override void OnActivityResult(int requestCode, Result resultCode, Intent data)
        {
            base.OnActivityResult(requestCode, resultCode, data);

            if (authContext != null)
            {
                authContext.OnActivityResult(requestCode, (int)resultCode, data);
            }
        }

        class TestCallback : Java.Lang.Object, IAuthenticationCallback
        {
            Context context;

            public TestCallback(Context ctx)
            {
                context = ctx;
            }

            public void OnError(Java.Lang.Exception exc)
            {
                AlertDialog.Builder builder = new AlertDialog.Builder(context);
                builder.SetTitle("Error");
                builder.SetMessage(exc.Message);
                builder.SetCancelable(false);
                builder.SetPositiveButton("OK", (sender, args) => { });
                builder.Create().Show();
            }

            public void OnSuccess(Java.Lang.Object result)
            {
                AuthenticationResult aresult = result.JavaCast<AuthenticationResult>();
                if (aresult != null)
                {
                    ((MainActivity)context).AadUserInfo = aresult.UserInfo;

                    ((MainActivity)context).GetContacts();
                }
            }
        }
    }
}

