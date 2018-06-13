using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using Windows.Storage;
using Microsoft.Identity.Client;

namespace NativeO365CalendarEvents
{
  public class AuthenticationHelper
  {
    static string clientId = App.Current.Resources["ida:ClientID"].ToString();
    public static string[] Scopes = { "User.Read", "Calendars.Read" };

    public static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId);

    public static string TokenForUser = null;
    public static DateTimeOffset Expiration;
    public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

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

  }

}
