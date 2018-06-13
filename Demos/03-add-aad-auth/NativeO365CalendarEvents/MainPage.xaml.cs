using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

using System.Threading.Tasks;
using Windows.Storage;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace NativeO365CalendarEvents
{
  /// <summary>
  /// An empty page that can be used on its own or navigated to within a Frame.
  /// </summary>
  public sealed partial class MainPage : Page
  {
    private bool _connected = false;
    private string _emailAddress = null;
    private string _displayName = null;
    public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

    public MainPage()
    {
      this.InitializeComponent();
    }

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

    private async Task<bool> SignInCurrentUserAsync()
    {
      var token = await AuthenticationHelper.GetTokenForUserAsync();
      if (token != null)
      {
        this._emailAddress = (string)_settings.Values["userEmail"];
        this._displayName = (string)_settings.Values["userName"];
        return true;
      }
      else
      {
        return false;
      }
    }

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
      }
      else
      {
        EventList.ItemsSource = null;
        AuthenticationHelper.SignOut();

        InfoText.Text = "Press the following button to connect to Office 365.";
        ConnectButton.Content = "Connect";

        _connected = false;
      }

      ProgressBar.Visibility = Visibility.Collapsed;
    }

    private async void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
    }

  }
}
