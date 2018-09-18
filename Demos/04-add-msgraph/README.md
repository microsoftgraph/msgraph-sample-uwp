# Integrate Microsoft Graph into the Application

In this demo you will incorporate the Microsoft Graph into the application. For this application, you will use the Microsoft Graph .NET SDK.

> This demo builds off the final product from the previous demo.

1. Create a new event model object to store the events:
    1. In the **Solution Explorer** tool window, right-click the project and select **Add > New Folder**.
        1. Name the folder **Models**.
    1. Right-click the **Models** folder and select **Add > Class**.
        1. Name the folder **CalendarEvent**.
    1. Add the following `using` statements to the end of the existing `using` statements in the **MainPage.xaml.cs** file:

        ```cs
        using System.ComponentModel.DataAnnotations;
        ```

    1. Change the accessor of the class to `public`.
    1. Add the three members to the class to store the subject and dates related to the event. The final code for the `CalendarEvent` class should be as follows:

        ```cs
        public class CalendarEvent
        {
          public string Subject { get; set; }

          [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
          public DateTimeOffset? Start { get; set; }

          [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
          public DateTimeOffset? End { get; set; }
        }
        ```

1. Extend the `AuthenticationHelper` class to return a `GraphServiceClient` object used to call the Microsoft Graph:
    1. Open the **AuthenticationHelper.cs** file.
    1. Add the following `using` statements to the end of the existing `using` statements in the **MainPage.xaml.cs** file:

        ```cs
        using Microsoft.Graph;
        ```

    1. Add the following code to the `AuthenticationHelper` class to create a create a new instance of the `GraphServiceClient`:

        ```cs
        private static GraphServiceClient graphClient = null;

        public static GraphServiceClient GetAuthenticatedClient()
        {
          AuthenticationHelper.graphClient = new GraphServiceClient(
            new DelegateAuthenticationProvider(async (requestMessage) =>
            {
              string accessToken = await AuthenticationHelper.GetTokenForUserAsync();
              requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            }));
          return graphClient;
        }
        ```

    1. Locate the `SignOut()` method and add the following code to the end of the method to destroy the authenticated `GraphServiceClient` object.

        ```cs
        AuthenticationHelper.graphClient = null;
        ```

1. Update the application to retrieve calendar events for the currently logged in user:
    1. Open the **MainPage.xaml.cs** file.
    1. Add the following `using` statements to the end of the existing `using` statements in the **MainPage.xaml.cs** file:

        ```cs
        using NativeO365CalendarEvents.Models;
        using Microsoft.Graph;
        ```

    1. Add the following method that will retrieve calendar events from your Office 365 calendar and bind them to the XAML list control:

        ```cs
        public async Task ReloadEvents()
        {
          var graphService = AuthenticationHelper.GetAuthenticatedClient();
          var request = graphService.Me.Events.Request(new Option[] {
            new QueryOption("top", "20"),
            new QueryOption("skip", "0")
          });
          var userEventsCollectionPage = await request.GetAsync();

          var calendarEvents = new List<CalendarEvent>();
          foreach (var calEvent in userEventsCollectionPage)
          {
            calendarEvents.Add(new CalendarEvent
            {
              Subject = !string.IsNullOrEmpty(calEvent.Subject)
                ? calEvent.Subject
                : string.Empty,
              Start = !string.IsNullOrEmpty(calEvent.Start.DateTime)
                ? DateTime.Parse(calEvent.Start.DateTime)
                : new DateTime(),
              End = !string.IsNullOrEmpty(calEvent.End.DateTime)
                ? DateTime.Parse(calEvent.End.DateTime)
                : new DateTime()
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
        ```

    1. Update the `ReloadButton_Click()` method by adding the following code to it to call the method you just added:

        ```cs
        ProgressBar.Visibility = Visibility.Visible;
        await ReloadEvents();
        ProgressBar.Visibility = Visibility.Collapsed;
        ```

    1. Test the application by pressing **F5**.
    1. After the app loads, press the **Connect** button to initiate the login process:
    1. After successfully signing in, you will see the application display the current name and email of the signed in user.

        Select the **Load Events** button.

        After a moment you should see a list of events from your Office 365 calendar.

        ![Screenshot showing the application displaying events from the user's calendar.](../../Images/vs-app-04.png)