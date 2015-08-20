# Office 365 Profile Angular sample

The Office 365 Profile Angular sample uses the Office 365 unified endpoint (preview) to get user profile data from various services such as Active Directory, SharePoint, and OneDrive.

The list of users in the Active Directory tenant appears in the sidebar. After selecting a user, their information, including user details, email address, hire date, manager, direct reports, and public files, are displayed in the main section of the app. All of this information comes from the unified endpoint.

![O365 Profile Angular sample screenshot](./README assets/screenshot.PNG)

<a name="prerequisites"></a>
## Prerequisites

This sample requires the following:
* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 
* An Office 365 account. You can sign up for [an Office 365 Developer subscription](http://aka.ms/ro9c62) that includes the resources that you need to start building Office 365 apps.
* A Microsoft Azure tenant to register your application. Azure Active Directory provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](http://aka.ms/jjm0q7).

**Note**  You will also need to ensure your Azure subscription is bound to your Office 365 tenant. Check out the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://aka.ms/lrb3ln) for instructions. In this post, the *Adding a new directory* section will explain how to do this. You can also read [Associate your Office 365 account with Azure AD to create and manage apps](http://aka.ms/fv273q) for more information.

<a name="configure"></a>
## Register and configure the app

1. Sign in to the [Azure Management Portal](https://manage.windowsazure.com/) using your Office 365 business account credentials.

2. Click the **Active Directory** node in the left column and select the directory linked to your Office 365 subscription.

3. Select the **Applications** tab and then **Add** at the bottom of the screen.

4. On the pop-up, select **Add an application my organization is developing**. Then click the arrow to continue. 

5. Choose a name for the app, such as *O365-Angular-Profile*, and select **Web application and/or web API** as its **Type**. Then click the arrow to continue.

6. The value of **Sign-on URL** is the URL where the application will be hosted. Use *http://127.0.0.1:8080/* for the sample project.

7. The value of **App ID URI** is a unique identifier for Azure AD to identify the app. You can use http://{your_subdomain}/O365-Angular-Profile, where {your_subdomain} is the subdomain of .onmicrosoft you specified while signing up for your Office 365 business account. Then click the check mark to provision the application.

8. Once the application is successfully added, the Quick Start page for the application appears. From here, select the **Configure** tab.

9. Scroll down to the **permissions to other applications** section and click the **Add application** button.

10. In this tutorial, we'll use the Office 365 unified API to get user data, so add the **Office 365 unified API (preview)** application. Click the plus sign in the application's row and then click the check mark at the top right to add it. Then click the check mark at the bottom right to continue.

11. In the **Office 365 unified API (preview)** row, select **Delegated Permissions**, and in the selection list, select **Read all users' basic profiles**, **Read items in all site collections**, **Read users' files** and **Access directory as the signed in user**.

12. Click **Save** to save the app's configuration.

### Configure the app to allow the OAuth 2.0 implicit grant flow

In order to get an access token for unified API requests, the application will use the OAuth implicit grant flow. You need to update the application's manifest to allow the OAuth implicit grant flow because it is not allowed by default. 

1. Select the **Configure** tab of the application's entry in the Azure Management Portal. 

2. Using the **Manage Manifest** button in the drawer, download the manifest file for the application and save it to the computer.

3. Open the manifest file with a text editor. Search for the *oauth2AllowImplicitFlow* property. By default it is set to *false*; change it to *true* and save the file.

4. Using the **Manage Manifest** button, upload the updated manifest file.

<a name="run"></a>
## Run the app

Open *app/scripts/config.js* and replace *{your_tenant}* with the subdomain of .onmicrosoft you specified for your Office 365 tenant and replace *{client_ID}* with the client ID of your registered Azure application (found on the **Configure** tab of your application's entry in the Azure Management Portal).

Next, install the necessary dependencies and run the project via the command line. Begin by opening a command prompt and navigating to the root folder. Once there, follow the steps below.

1. Install project dependencies by running ```npm install```.
2. Now that all the project dependencies are installed, start the development server by running ```node server.js``` in the root folder.
3. Navigate to ```http://127.0.0.1:8080/``` in your web browser.

## Security notice
[ADAL JS](https://github.com/AzureAD/azure-activedirectory-library-for-js) does not validate the token received from Azure AD. It relies on the app’s backend to do so, and until you call the backend, you don’t know if the user obtained an acceptable token. Business applications should have a server-side component for user authentication built into the web application for security reasons. Without this backend token validation, your app is susceptible to security attacks such as the [confused deputy problem](https://en.wikipedia.org/wiki/Confused_deputy_problem). Check out this [blog post](http://www.cloudidentity.com/blog/2015/02/19/introducing-adal-js-v1/) for more information.

<a name="questions-and-comments"></a>
## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/O365-Angular-Profile/issues).
- For general questions about the Office 365 APIs, post to [Stack Overflow](http://stackoverflow.com/). Make sure that your questions or comments are tagged with [office365].
  
<a name="additional-resources"></a>
## Additional resources

* [Office Dev Center](http://dev.office.com/)
* [Office 365 APIs platform overview](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [Office 365 unified API overview](https://msdn.microsoft.com/en-us/office/office365/howto/office-365-unified-api-overview)
* [Create an Angular app with Office 365 APIs](http://aka.ms/get-started-with-js)
* [Other Office 365 Profile samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Profile)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
