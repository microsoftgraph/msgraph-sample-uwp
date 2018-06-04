# Extend the UWP app for Azure AD Authentication

In this demo you will extend the application from the previous demo to support authentication with Azure AD. This is required to obtain the necessary OAuth access token to call the Microsoft Graph. In this step you will integrate the Microsoft Authentication Library (MSAL) into the application.

> This demo builds off the final product from the previous demo.

1. Open the **App.xaml** file.
1. Add the following markup to the `<Application>` element. This will specify two formatting options as well as define the application ID and MSAL redirect URLs:

    ```xml
    <Application.Resources>
      <SolidColorBrush x:Key="SampleHeaderBrush" Color="#007ACC" />
      <SolidColorBrush x:Key="ApplicationPageBackgroundThemeBrush" Color="#1799F0" />

      <x:String x:Key="ida:ClientID">ENTER_APP_ID</x:String>
      <x:String x:Key="ida:ReturnUrl">ENTER_APP_CUSTOM_REDIRECT_URI</x:String>
    </Application.Resources>
    ```

1. Update the `ida:ClientID` & `ida:ReturnUrl` values to those you copied when creating a the Azure AD application in a previous demo.
1. Add an authentication helper class:
    1. Right-click the project in the **Solution Explorer** tool window and select **Add > Class**.
    1. Set the **Name** of the class to **AuthenticationHelper.cs** and select **Add**.
    1. In the **AuthenticationHelper.cs** file, add the following `using` statements after the default statements added:

        ```cs
        using System.Net.Http.Headers;
        using Windows.Storage;
        using Microsoft.Identity.Client;
        ```

    1. Update the accessor of the `AuthenticationHelper` class to be public by adding the `public` keyword:

        ```cs
        public class AuthenticationHelper
        ```

    1. Add the following members to the `AuthenticationHelper` class that will be used throughout this class:

        ```cs
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();
        public static string[] Scopes = { "User.Read", "Calendars.Read" };

        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);

        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;
        ```

    1. Add the following method `GetTokenForUserAsync()` that will start the interactive authentication process with the current user of the application. If the authentication process is successful, the user's name and email address are stored in Windows roaming storage and the token is returned to the caller:

        ```cs
        internal static async Task<string> GetTokenForUserAsync()
        {
          AuthenticationResult authResult;
          try
          {
            authResult = await IdentityClientApp.AcquireTokenSilentAsync(Scopes, IdentityClientApp.Users.First());
            TokenForUser = authResult.AccessToken;

            _settings.Values["userEmail"] = authResult.User.DisplayableId;
            _settings.Values["userName"] = authResult.User.Name;
          }

          catch (Exception)
          {
            if (TokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
            {
              authResult = await IdentityClientApp.AcquireTokenAsync(Scopes);

              TokenForUser = authResult.AccessToken;
              Expiration = authResult.ExpiresOn;

              _settings.Values["userEmail"] = authResult.User.DisplayableId;
              _settings.Values["userName"] = authResult.User.Name;
            }
          }

          return TokenForUser;
        }
        ```

    1. Add the following method `SignOut()` to log the user out:

        ```cs
        public static void SignOut()
        {
          foreach (var user in IdentityClientApp.Users)
          {
            IdentityClientApp.Remove(user);
          }

          TokenForUser = null;

          _settings.Values["userID"] = null;
          _settings.Values["userEmail"] = null;
          _settings.Values["userName"] = null;
        }
        ```

1. Update the application's user interface to leverage the Azure AD authentication code you just added:
    1. In the **Solution Explorer** tool window, expand the **MainPage.xaml** file and double-click the **MainPage.xaml.cs** file:

        ![Screenshot showing the MainPage.xaml.cs file in the Solution Explorer tool window](../../Images/vs-code-project-01.png)

    1. Add the following `using` statements to the end of the existing `using` statements in t he **MainPage.xaml.cs** file:

        ```cs
        using System.Threading.Tasks;
        using Windows.Storage;
        ```

    1. Add the following members to the `MainPage` partial class

        ```cs
        private bool _connected = false;
        private string _emailAddress = null;
        private string _displayName = null;
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;
        ```

    1. Add the following method that will run when the application is initialized:

        ```cs
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
          if (!App.Current.Resources.ContainsKey("ida:ClientID"))
          {
            InfoText.Text = "Oops - It looks like this app is not registered with Office 365, because we don't see a client id in App.xaml.";
            ConnectButton.IsEnabled = false;
          }
          else
          {
            InfoText.Text = "Press the following button to connect to Office 365.";
            ConnectButton.IsEnabled = true;
          }
        }
        ```

    1. Add the following method that will be used to trigger the authentication process:

        ```cs
        private async Task<bool> SignInCurrentUserAsync()
        {
          var token = await AuthenticationHelper.GetTokenForUserAsync();
          if (token != null)
          {
            Debug.WriteLine("Token: " + token);
            this._emailAddress = (string)_settings.Values["userEmail"];
            this._displayName = (string)_settings.Values["userName"];
            return true;
          }
          else
          {
            return false;
          }
        }
        ```

    1. Add the following method that acts as the event handler when the **Connect** button is pressed. It will start the signin / signout process depending on the current logged in state.

        ```cs
        private async void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
          ProgressBar.Visibility = Visibility.Visible;

          if (!_connected)
          {
            if (await SignInCurrentUserAsync())
            {
              InfoText.Text = "Hi " + _displayName + " (" + _emailAddress + ")!";
              ConnectButton.Content = "Disconnect";
              _connected = true;
            }
            else
            {
              InfoText.Text = "Oops! We couldn't connect to Office 365. Check your debug output for errors.";
            }
          } else
          {
            EventList.ItemsSource = null;
            AuthenticationHelper.SignOut();

            InfoText.Text = "Press the following button to connect to Office 365.";
            ConnectButton.Content = "Connect";

            _connected = false;
          }

          ProgressBar.Visibility = Visibility.Collapsed;
        }
        ```

    1. Add the following method stub that you will implement later. It is used by the button in the UI that will retrieve events from your Office 365 calendar using the Microsoft Graph:

        ```cs
        private async void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
        }
        ```

    1. Test the application by pressing **F5**.

        ![Screenshot of the UWP application running in debug mode](../../Images/vs-app-01.png)

    1. After the app loads, press the **Connect** button to initiate the login process:

        ![Screenshot of the Azure AD login prompt](../../Images/vs-app-02.png)

    1. After successfully logging in, you may be prompted to consent to the permissions requested by the application. If prompted, agree to the consent dialog.

    1. After successfully signing in, you will see the application display the current name and email of the signed in user:

        ![Screenshot showing the application after successfully signing into Azure AD](../../Images/vs-app-03.png)

        Now that the application is working with Azure AD, the next step is to implement the Microsoft Graph integration.

    1. Select the **Disconnect** button in the application and close the application.