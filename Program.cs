Console.WriteLine("Hi, Welcome to our Interview Scheduler");
Console.WriteLine("Make sure you have a CSV file with the following columns: Name, Email, InterviewTime, Interviewer, InterviewerMail");
Console.WriteLine("Just log in to your Microsoft account and we will take care of the rest!");

var settings = Settings.LoadSettings();

// Initialize Graph
InitializeGraph(settings);

// Greet the user by name
await GreetUserAsync();

int choice = -1;

while (choice != 0)
{
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Schedule Interviews");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (System.FormatException)
    {
        choice = -1;
    }

    switch (choice)
    {
        case 0:
            Console.WriteLine("Thank you for using our Interview Scheduler!");
            break;
        case 1:
            await SendInterviewInvitationsAsync();
            break;
        default:
            Console.WriteLine("Invalid choice! Please try again.");
            break;
    }
}

void InitializeGraph(Settings settings)
{
    GraphHelper.InitializeGraphForUserAuth(settings,
        (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });
}

async Task GreetUserAsync()
{
    try
    {
        var user = await GraphHelper.GetUserAsync();
        Console.WriteLine($"Hello, {user?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName ?? ""}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error getting user: {ex.Message}");
    }
}

async Task SendInterviewInvitationsAsync()
{
    try
    {
        await GraphHelper.SendInterviewInvitationsAsync();
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error while creating a meet: {ex.Message}");
    }
}

