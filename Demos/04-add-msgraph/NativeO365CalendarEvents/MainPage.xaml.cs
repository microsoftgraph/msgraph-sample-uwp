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
using NativeO365CalendarEvents.Models;
using Microsoft.Graph;


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

    public async Task ReloadEvents()
    {
      var graphService = AuthenticationHelper.GetAuthenticatedClient();
      var request = graphService.Me.Events.Request(new Option[] { new QueryOption("top", "20"), new QueryOption("skip", "0") });
      var userEventsCollectionPage = await request.GetAsync();

      var calendarEvents = new List<CalendarEvent>();
      foreach (var calEvent in userEventsCollectionPage)
      {
        calendarEvents.Add(new CalendarEvent
        {
          Subject = !string.IsNullOrEmpty(calEvent.Subject) ? calEvent.Subject : string.Empty,
          Start = !string.IsNullOrEmpty(calEvent.Start.DateTime) ? DateTime.Parse(calEvent.Start.DateTime) : new DateTime(),
          End = !string.IsNullOrEmpty(calEvent.End.DateTime) ? DateTime.Parse(calEvent.End.DateTime) : new DateTime()
        });
      }

      if (calendarEvents != null && calendarEvents.Count > 0)
      {
        EventList.ItemsSource = calendarEvents;
        EventList.Visibility = Visibility.Visible;
      }
      else
      {
        EventList.ItemsSource = null;
        EventList.Visibility = Visibility.Collapsed;
      }
    }

    private async void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
      ProgressBar.Visibility = Visibility.Visible;
      await ReloadEvents();
      ProgressBar.Visibility = Visibility.Collapsed;
    }

  }
}
