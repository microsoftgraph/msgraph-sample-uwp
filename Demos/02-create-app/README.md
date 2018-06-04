# Create a UWP application with Visual Studio 2017

In this demo you will create a UWP application with Visual Studio and wire up a few events. This will be used as the starting point prior to creating the Microsoft Graph enabled application.

1. Open Visual Studio 2017.
1. In Visual Studio, select **File > New > Project**.
1. In the **New Project** dialog, do the following:
    1. Select **Templates > Windows Universal**.
    1. Select **Blank App (Universal Windows)**.
    1. Enter **NativeO365CalendarEvents** for the **Name** of the project.

          ![Screenshot of creating a new Windows Universal project](../../Images/vs-create-project-01.png)

    1. In the **New Universal Windows Platform Project** dialog, you can select anything you like as nothing in this lab depends on specific Windows features.

        Select **OK** after specifying your desired **Target version** & **Minimum version**.

        ![Screenshot of creating a new Windows Universal project](../../Images/vs-create-project-02.png)

1. Add the necessary NuGet Packages to the project:
    1. In the **Solution Explorer** tool window, right-click the **References** node and select **Manage NuGet Packages...**:

        ![Screenshot selecting Manage NuGet Packages in Visual Studio](../../Images/vs-setup-project-01.png)

    1. Add the Microsoft Graph .NET SDK to the project:
        1. Select the **Browse** tab and enter **Microsoft Graph** in the search box.git push
        1. Select the **Microsoft.Graph** client in the results.
        1. Select **Install** to install the package.

            ![Screenshot installing the Microsoft Graph .NET SDK NuGet package](../../Images/vs-setup-project-02.png)

            If prompted, accept all licenses.

    1. Add the Microsoft Authentication Library (MSAL) Preview to the project:
        1. Select the **Browse** tab and enter **Microsoft.Identity.Client** in the search box.
        1. Select the **Include Prerelease** checkbox to include libraries currently in preview.

            > MSAL is currently in preview at the time of writing.

        1. Select the **Microsoft.Identity.Client** client in the results.
        1. Select **Install** to install the package.

            ![Screenshot installing the MSAL .NET SDK NuGet package](../../Images/vs-setup-project-03.png)

            If prompted, accept all licenses.

## Create the Application User Interface

The first step is to create the shell of the user experience; creating a workable storyboard.

1. Open the **MainPage.xaml** file.
1. Replace the empty `<Grid></Grid>` XAML element with the following XAML.

    This XAML will create an interface broken into two sections:

      * **Header**:
        * A button to sign in & sign out of your Office 365 account
        * A button to refresh the list of events from your calendar
        * Two (2) text blocks to display the application title and name of the signed in user.
      * **Body**:
        * Progress animation control shown when authenticating with Azure AD & loading events from Office 365 using the Microsoft Graph
        * List control showing events from your calendar

    ```xml
    <Grid Background="{StaticResource ApplicationPageBackgroundThemeBrush}">
      <Grid x:Name="MainGrid">
        <Grid.RowDefinitions>
          <RowDefinition Height="90"/>
          <RowDefinition Height="90"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Button Grid.Row="0" x:Name="ConnectButton" HorizontalAlignment="Right" Content="Connect" FontSize="14.667" Click="ConnectButton_Click" Height="40" Width="100" Background="Transparent" FontWeight="SemiBold" Foreground="White" BorderBrush="White"/>
        <TextBlock Grid.Row="1" Grid.Column="1"  x:Name="InfoText" Foreground="White" VerticalAlignment="Center" HorizontalAlignment="Left" FontSize="{ThemeResource TextStyleLargeFontSize}" TextWrapping="Wrap" Width="750" Margin="0,10"/>
        <Button Grid.Row="1" x:Name="ReloadButton"  Margin="0,20,0,0" HorizontalAlignment="Center" VerticalAlignment="Top" Content="Load Events" FontSize="14.667"  Click="ReloadButton_Click" Height="40" Width="100" Background="Transparent" FontWeight="SemiBold" Foreground="White" BorderBrush="White"/>
        <TextBlock x:Name="appTitle" Grid.Row="0" Text="Windows Office 365 Calendar Events" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontSize="36" Style="{StaticResource HeaderTextBlockStyle}" Foreground="White" />

        <ProgressRing x:Name="ProgressBar" Visibility="Collapsed" Grid.Row="2" IsActive="True" Width="50" Height="50" Foreground="White" />
        <ListView x:Name="EventList" AutomationProperties.AutomationId="EventListView" AutomationProperties.Name="Items" TabIndex="1" Grid.Row="2" Margin="50,0,0,0" Padding="120,0,0,60" IsSwipeEnabled="False" SelectionMode="Single">
          <ListView.ItemTemplate>
            <DataTemplate>
              <Grid Margin="6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Column="1" Grid.Row="1" Margin="0,0,5,0">
                  <TextBlock Style="{StaticResource BodyTextBlockStyle}" Text="{Binding Subject}" TextWrapping="NoWrap" MaxHeight="40"  FontSize="20" Foreground="White"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Grid.Row="2"  Margin="0,0,5,0">
                  <TextBlock Style="{StaticResource BodyTextBlockStyle}" Text="Start Time: " TextWrapping="NoWrap" MaxHeight="40" FontSize="16"  Foreground="White"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Grid.Row="2"  Margin="80,0,5,0">
                  <TextBlock Style="{StaticResource BodyTextBlockStyle}" Text="{Binding Start}" TextWrapping="NoWrap" MaxHeight="40" FontSize="16"  Foreground="White"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Grid.Row="3"  Margin="0,0,5,0">
                  <TextBlock Style="{StaticResource BodyTextBlockStyle}" Text="End Time: " TextWrapping="NoWrap" MaxHeight="40" FontSize="16"  Foreground="White"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Grid.Row="3" Margin="80,0,5,0">
                  <TextBlock Style="{StaticResource BodyTextBlockStyle}" Text="{Binding End}" TextWrapping="NoWrap" MaxHeight="40" FontSize="16"  Foreground="White"/>
                </StackPanel>
              </Grid>
            </DataTemplate>
          </ListView.ItemTemplate>
          <ListView.ItemContainerStyle>
            <Style TargetType="FrameworkElement">
              <Setter Property="Margin" Value="0,0,0,10"/>
            </Style>
          </ListView.ItemContainerStyle>
        </ListView>
      </Grid>

    </Grid>
    ```