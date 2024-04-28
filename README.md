# Microsoft Graph .NET Client Library For Mail

The Microsoft Graph .NET Mail Client Library targets .Net 6.0.

Integrate the [Microsoft Graph API](https://graph.microsoft.com) into your .NET project!
## Installation via NuGet

To install the client library via NuGet:

* Search for `Microsoft.Graph.Client.API` in the NuGet Library, or
* Type `Install-Package Microsoft.Graph.Client.API` into the Package Manager Console.

## Getting started

### 1. Register your application

Register your application to use Microsoft Graph API using the [Microsoft Application Registration Portal](https://aka.ms/appregistrations).

### 2. Authenticate for the Microsoft Graph service

The Microsoft Graph .NET Client Library supports the use of TokenCredential classes in the [Azure.Identity](https://www.nuget.org/packages/Azure.Identity) library.

You can read more about available Credential classes [here](https://docs.microsoft.com/en-us/dotnet/api/overview/azure/identity-readme#key-concepts) and examples on how to quickly setup TokenCredential instances can be found [here](https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/tokencredentials.md).

The recommended library for authenticating against Microsoft Identity (Azure AD) is [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet).

For an example of authenticating a UWP app using the V2 Authentication Endpoint, see the [Microsoft Graph UWP Connect Library](https://github.com/OfficeDev/Microsoft-Graph-UWP-Connect-Library).

### 3. Create a Microsoft Graph client object with an authentication provider

An instance of the **GraphServiceClient** class handles building requests,
sending them to Microsoft Graph API, and processing the responses. To create a
new instance of this class, you need to provide an instance of
`IAuthenticationProvider` which can authenticate requests to Microsoft Graph.

### 4. Make requests to the graph mail

Once you have completed authentication, you can
begin to make calls to the service. The requests in the SDK follow the format
of the Microsoft Graph Mail API's RESTful syntax.
