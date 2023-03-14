using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using CsvHelper;
using System.Globalization;

class GraphHelper
{
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static async Task<string> GetUserTokenAsync()
{
    // Ensure credential isn't null
    _ = _deviceCodeCredential ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    // Ensure scopes isn't null
    _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

    // Request token with given scopes
    var context = new TokenRequestContext(_settings.GraphUserScopes);
    var response = await _deviceCodeCredential.GetTokenAsync(context);
    return response.Token;
}

public static async Task SendMailAsync(string subject, ItemBody body, string recipient)
{
    // Ensure client isn't null
    _ = _userClient ??
        throw new System.NullReferenceException("Graph has not been initialized for user auth");

    // Create a new message
    var message = new Message
    {
        Subject = subject,
        Body = body,
        ToRecipients = new Recipient[]
        {
            new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }
        }
    };

    // Send the message
    await _userClient.Me
        .SendMail(message)
        .Request()
        .PostAsync();
}

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.TenantId, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }

    public static Task<User?> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
        .Request()
        .Select(u => new
        {
            // Only request specific properties
            u.DisplayName,
            u.Mail,
            u.UserPrincipalName
        })
        .GetAsync();
    }

    public async static Task<string?> CreateTeamsMeet(DateTime startDateTime)
    {
        _ = _userClient ??
           throw new System.NullReferenceException("Graph has not been initialized for user auth");

        var sub = "Interview Meeting";
        var srtTime = startDateTime;
        var hrs = "1";

        var requestBody = new OnlineMeeting
        {
            StartDateTime = srtTime,
            EndDateTime = srtTime.AddHours(Int32.Parse(hrs)),
            Subject = sub
        };

        var meetLink = await _userClient.Me.OnlineMeetings.Request().AddAsync(requestBody);
      
        return meetLink?.JoinWebUrl;
    }

    public async static Task SendInterviewInvitationsAsync()
  {
      try
      {
          var settings = Settings.LoadSettings();

          // Read data from CSV file
          Console.WriteLine("Please enter the path to the CSV file:");
          var csvFilePath = Console.ReadLine();
          var recipients = ReadRecipientsFromCsv(csvFilePath);

          foreach (var recipient in recipients)
        {
            Console.WriteLine($"Sending email to {recipient.Name} ({recipient.Email})...");
            
            // Create Teams meeting
            var meetingURL = await GraphHelper.CreateTeamsMeet(recipient.InterviewTime);

            // Send email to interviewee
            var message = new Message
            {
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient.Email,
                            Name = recipient.Name
                        }
                    }
                },
                Subject = "Interview invitation",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"Dear {recipient.Name},<br><br>You have been invited for an interview on {recipient.InterviewTime}. Please click <a href='{meetingURL}'>here</a> to join the Teams meeting.<br><br>Best regards,<br>The Interview Team"
                }
            };

            await SendMailAsync(message.Subject, message.Body, recipient.Email);
            Console.WriteLine("Email sent to interviewee!\n");

            // Send email to interviewer
            Console.WriteLine($"Sending email to {recipient.Interviewer} ({recipient.InterviewerMail})...");
            
            message = new Message
            {
                ToRecipients = new List<Recipient>
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = recipient.InterviewerMail,
                            Name = recipient.Interviewer
                        }
                    }
                },
                Subject = "Interview Scheduled",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"Dear {recipient.Interviewer},<br><br>You have been assigned to take an interview on {recipient.InterviewTime}. <br> Candidate Name is {recipient.Name}. Please click <a href='{meetingURL}'>here</a> to join the Teams meeting.<br><br>Best regards,<br>The Interview Team"
                }
            };
            await SendMailAsync(message.Subject, message.Body, recipient.Email);
            Console.WriteLine("Email sent to interviewer!\n");

            Console.WriteLine("Interview scheduled!\n\n");
        }

        Console.WriteLine("All emails have been sent. Thank you for using this tool!\n\n\n");
      }
      catch (Exception ex)
      {
          Console.WriteLine($"An error occurred: {ex.Message}");
      }
  }

static List<RecipientData> ReadRecipientsFromCsv(string filePath)
    {
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        return csv.GetRecords<RecipientData>().ToList();
    }

class RecipientData
{
    public string Name { get; set; }
    public string Email { get; set; }
    public DateTime InterviewTime { get; set; }
    public string Interviewer { get; set; }
    public string InterviewerMail { get; set; }
}
}