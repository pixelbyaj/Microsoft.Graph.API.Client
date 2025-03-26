# Microsoft Graph .NET Client Library For Mail

[![NuGet Version](https://img.shields.io/nuget/v/MSGraph.Mail.Client)](https://www.nuget.org/packages/MSGraph.Mail.Client)
[![NuGet Downloads](https://img.shields.io/nuget/dt/MSGraph.Mail.Client)](https://www.nuget.org/packages/MSGraph.Mail.Client)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue)](https://github.com/pixelbyaj/Microsoft.Graph.API.Client/blob/main/LICENSE)

The Microsoft Graph .NET Mail Client Library targets .NET 6.0.

Integrate the [Microsoft Graph API](https://graph.microsoft.com) into your .NET project!

## Installation via NuGet


Search for `MSGraph.Mail.Client` in the NuGet Library

OR

Type write below into the Package Manager Console.
```cli
Install-Package MSGraph.Mail.Client
```

## Getting started

### 1. Register your application

1. Open a browser and navigate to the [Microsoft Entra admin center](https://entra.microsoft.com/) and log in using a Work or School Account.

2. Register your application to use Microsoft Graph API using the [Microsoft Application Registration Portal](https://aka.ms/appregistrations).

3. Select **New registration**. Enter a name for your application.

4. Set **Supported account types** as desired.

5. If you want to set up your application as a background service, leave the Redirect URI empty. Otherwise, if you want to set it up as a delegate (sign-in) user, add **http://localhost**.

6. Select **Register**. On the application's Overview page, copy the value of the **Application (client) ID** and save it. You will need it in the next step. If you chose **Accounts in this organizational directory only** for Supported account types, also copy the **Directory (tenant) ID** and save it.

![Application Registration](https://raw.githubusercontent.com/pixelbyaj/Microsoft.Graph.API.Client/refs/heads/main/assets/image-1.png)

**Note:** If you want to use it for your personal email account like outlook.com or live.in, use `tenantId` as **consumer**.

7. Select **Authentication** under Manage. Locate the **Advanced settings** section and change the **Allow public client flows** toggle to **Yes**, then choose **Save**.

![Allow Public Client Flows](https://raw.githubusercontent.com/pixelbyaj/Microsoft.Graph.API.Client/refs/heads/main/assets/image-1.png)

### 2. Create a Microsoft Graph mail client object with an authentication provider

#### Interactive Flow
```csharp
Settings settings = new();
settings.ClientId = "";
settings.TenantId = "";
// We need to pass the scopes which we require and the same has been set at the API Permission in Azure
settings.GraphUserScopes = new string[] {
  "openid",
  "profile",
  "offline_access",
  "user.read",
  "mail.readbasic",
  "mail.read",
  "mail.send"
};

IAuthenticationProvider authenticationProvider = new AuthenticationInteractiveProvider(settings);
```

#### Daemon or Client Secret Flow
```csharp
Settings settings = new();
settings.ClientId = "";
settings.TenantId = "";
settings.SecretId = "";

// We set the scope only in the API permission level in Azure.
IAuthenticationProvider authenticationProvider = new AuthenticationClientSecretProvider(settings);
```

### 3. Make requests to the graph mail

Once you have completed authentication, you can begin to make calls to the Email Graph service. The requests in the SDK follow the format of the Microsoft Graph Mail API's RESTful syntax.

```csharp
IEmailGraphService emailGraphService = new EmailGraphService(authenticationProvider);
```

### 4. Read and Send Emails

The `IEmailGraphService` interface provides three different APIs:

#### Get Emails
```csharp
Task<IList<EmailMessage>> GetEmailsAsync(
    int top = 10, 
    int limit = 10, 
    EmailRequestParameterInformation? requestInformation = null, 
    bool markRead = false
);
```
- **top**: Set the top parameter to fetch only that many emails. By default, it fetches the top 10 emails in descending order by `receivedDateTime`.
- **limit**: Limit the number of emails fetched. Set `limit` to `-1` to fetch all emails.
- **requestInformation**: Helps set `$filter`, `$search`, order by given fields, and includes attachments with emails.

#### Get Email Attachments
```csharp
Task<IList<EmailFileAttachment>> GetEmailAttachments(string? messageId);
```
- **messageId**: The email message ID received from the Graph API.

#### Send Email
```csharp
Task SendEmail(EmailMessage emailMessage);
```
- **emailMessage**: The email message object.

### Example

```csharp
await emailGraphService.SendEmail(new EmailMessage
{
    ToRecipients = new[] { "example@example.com" },
    Subject = "Test Email",
    BodyType = EmailBodyType.Html,
    BodyContent = "<h1>Hello, World!</h1>"
});
```

### Note

Please check the `EmailDemoApp` project for detailed examples.