// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=234238

namespace GraphTutorial
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class HomePage : Page
    {
        // <ConstructorSnippet>
        public HomePage()
        {
            this.InitializeComponent();

            if ((App.Current as App).IsAuthenticated)
            {
                HomePageMessage.Text = "Welcome! Please use the menu to the left to select a view.";
            }
        }
        // </ConstructorSnippet>
    }
}
