// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using CommunityToolkit.Authentication;
using CommunityToolkit.Graph.Extensions;
using Microsoft.Identity.Client;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace GraphTutorial
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        // <ConstructorSnippet>
        public MainPage()
        {
            this.InitializeComponent();

            // Load OAuth settings
            var oauthSettings = Windows.ApplicationModel.Resources.ResourceLoader.GetForCurrentView("OAuth");
            var appId = oauthSettings.GetString("AppId");
            var scopes = oauthSettings.GetString("Scopes");

            if (string.IsNullOrEmpty(appId) || string.IsNullOrEmpty(scopes))
            {
                Notification.Show("Could not load OAuth Settings from resource file.");
            }
            else
            {
                // Configure MSAL provider
                var msalClient = PublicClientApplicationBuilder.Create(appId)
                    .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
                    .Build();
                ProviderManager.Instance.GlobalProvider = new MsalProvider(msalClient, scopes.Split(' '));

                // Handle auth state change
                ProviderManager.Instance.ProviderStateChanged += ProviderUpdated;

                // Navigate to HomePage.xaml
                RootFrame.Navigate(typeof(HomePage));
            }
        }
        // </ConstructorSnippet>

        // <ProviderUpdatedSnippet>
        private void ProviderUpdated(object sender, ProviderStateChangedEventArgs e)
        {
            var globalProvider = ProviderManager.Instance.GlobalProvider;
            SetAuthState(globalProvider != null && globalProvider.State == ProviderState.SignedIn);
            RootFrame.Navigate(typeof(HomePage));
        }
        // </ProviderUpdatedSnippet>

        // <SetAuthStateSnippet>
        private void SetAuthState(bool isAuthenticated)
        {
            (Application.Current as App).IsAuthenticated = isAuthenticated;

            // Toggle controls that require auth
            Calendar.IsEnabled = isAuthenticated;
            NewEvent.IsEnabled = isAuthenticated;
        }
        // </SetAuthStateSnippet>

        private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
        {
            var invokedItem = args.InvokedItem as string;

            // <SwitchStatementSnippet>
            switch (invokedItem.ToLower())
            {
                case "new event":
                    RootFrame.Navigate(typeof(NewEventPage));
                    break;
                case "calendar":
                    RootFrame.Navigate(typeof(CalendarPage));
                    break;
                case "home":
                default:
                    RootFrame.Navigate(typeof(HomePage));
                    break;
            }
            // </SwitchStatementSnippet>
        }
    }
}
