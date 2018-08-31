# Create an Azure AD native application with the App Registration Portal

In this demo you will create a new Azure AD native application using the App Registry Portal (ARP).

1. Open a browser and navigate to the **App Registry Portal**: **apps.dev.microsoft.com** and login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.
1. Select **Add an app** at the top of the page.
1. On the **Register your application** page, set the **Application Name** to **NativeO365CalendarEvents** and select **Create**.

    ![Screenshot of creating a new app in the App Registration Portal website](../../Images/arp-create-app-01.png)

1. On the **NativeO365CalendarEvents Registration** page, under the **Properties** section, copy the **Application Id** Guid as you will need it later.

    ![Screenshot of newly created application's ID](../../Images/arp-create-app-02.png)

1. Scroll down to the **Platforms** section.

    1. Select **Add Platform**.
    1. In the **Add Platform** dialog, select **Native Application**.

        ![Screenshot creating a platform for the app](../../Images/arp-create-app-03.png)

    1. After the native application platform is created, copy the **Custom Redirect URIs** as you will need it later.

        ![Screenshot of the custom application URI for the native application](../../Images/arp-create-app-04.png)

        > Unlike application secrets that are only displayed a single time when they are created, the custom redirect URIs are always shown so you can come back and get this string if you need it later.

1. In the **Microsoft Graph Permissions** section, select **Add** next to the **Delegated Permissions** subsection.

    ![Screenshot of the Add button for adding a delegated permission](../../Images/arp-add-permission-01.png)

    In the **Select Permission** dialog, locate and select the permission **Calendars.Read** and select **OK**:

      ![Screenshot of adding the Calendars.Read permission](../../Images/arp-add-permission-02.png)

      ![Screenshot of the newly added Calendars.Read permission](../../Images/arp-add-permission-03.png)

1. Scroll to the bottom of the page and select **Save**.