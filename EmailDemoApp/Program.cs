// See https://aka.ms/new-console-template for more information
using Microsoft.Graph.Mail.Client;
using Microsoft.Graph.Mail.Client.Authentications;
using Microsoft.Graph.Mail.Client.Models;

var settings = EmailApp.Settings.LoadSettings();


IAuthenticationProvider authenticationProvider = new AuthenticationInteractiveProvider(settings);
//IAuthenticationProvider authenticationProvider = new AuthenticationClientSecretProvider(settings);
var user = await authenticationProvider.GetUserProfile(null);

Console.WriteLine($"Hello, {user?.UserPrincipalName}");

IEmailGraphService emailGraphService = new EmailGraphService(authenticationProvider);
int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. List my inbox");
    Console.WriteLine("2. Send mail");
    Console.WriteLine("3. Make a Graph call");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            // Exit the program
            Console.WriteLine("Goodbye...");
            break;
        case 1:
            // List emails from user's inbox
            await ListInboxAsync();
            break;
        case 2:
            // Send an email message
            await SendMailAsync();
            break;
        case 3:
            // Run any Graph code
            await MakeGraphCallAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

Task MakeGraphCallAsync()
{
    throw new NotImplementedException();
}

async Task SendMailAsync()
{
    await emailGraphService.SendEmail(new EmailMessage
    {
        ToRecipients = new[]
        {
            "abhishek2185@gmail.com"
        },
        CcRecipients = new[]
        {
            "pixelbyaj@gmail.com"
        },
        Subject = "Test Email Message",
        BodyType = EmailBodyType.Html,
        BodyContent = "<h1>Hello, Email 1</h1>",
        HasAttachments = true,
        FileAttachments = new List<EmailFileAttachment>
        {
            new EmailFileAttachment
            {
                ContentType = "text/xml",
                Name = "camt.053_test_1.xml",
                ContentBytes = File.ReadAllBytes("C:\\source\\data\\camt.053_test_1.xml")
            }
        }
    });
}

async Task ListInboxAsync()
{
    var messages = await emailGraphService.GetEmailsAsync(2, -1, new Microsoft.Graph.Mail.Client.Models.EmailRequestParameterInformation
    {
        IsRead = false,
        IncludeAttachments = true,
    });
    foreach (var message in messages)
    {
        Console.WriteLine(message.From);
        Console.WriteLine(string.Join(", ", message.ToRecipients));
        Console.WriteLine(message.Subject);
        Console.WriteLine(message.BodyPreview);
        if (message.HasAttachments == true)
        {
            if (message.FileAttachments != null)
            {
                foreach (var attachment in message.FileAttachments)
                {
                    Console.WriteLine(attachment.Name);
                    Console.WriteLine(attachment.IsInline);
                }
            }
        }
    }
}