using Microsoft.Identity.Client;

namespace JunkEmailCleaner;

internal static class Application
{
    private const string ClientIdEnvironmentVariable = "JUNK_EMAIL_CLEANER_CLIENT_ID";

    public static async Task<int> Run(string[] args)
    {
        try
        {
            return await RunApplication(args);
        }
        catch (MsalUiRequiredException exception)
        {
            ConsoleLog.Error(
                $"JunkEmailCleaner failed because Microsoft requires another interactive Outlook login ({exception.ErrorCode}): {exception.Message} Run this program once with --login, then retry.",
                exception);
            return 3;
        }
        catch (Exception exception)
        {
            ConsoleLog.Error("JunkEmailCleaner failed. The exception and its underlying causes follow.", exception);
            return 1;
        }
    }

    private static async Task<int> RunApplication(string[] args)
    {
        ConsoleLog.Log($"JunkEmailCleaner started as operating-system user {Environment.UserName}.");

        if (args.Contains("--help", StringComparer.OrdinalIgnoreCase))
        {
            ConsoleLog.Log("Help was requested; no authentication or mailbox operations will run.");
            ShowHelp();
            ConsoleLog.Log("Help completed successfully.");
            return 0;
        }

        var clientId = Environment.GetEnvironmentVariable(ClientIdEnvironmentVariable);

        if (string.IsNullOrWhiteSpace(clientId))
        {
            ConsoleLog.Error(
                $"Configuration failed because {ClientIdEnvironmentVariable} is missing or empty. Set it to the Entra application client ID.");
            return 2;
        }

        ConsoleLog.Log($"Found a nonempty {ClientIdEnvironmentVariable} configuration value.");

        var login = args.Contains("--login", StringComparer.OrdinalIgnoreCase);
        var dryRun = args.Contains("--dry-run", StringComparer.OrdinalIgnoreCase);

        if (login)
        {
            ConsoleLog.Log("Selected interactive login mode.");
        }
        else if (dryRun)
        {
            ConsoleLog.Log("Selected dry-run mode.");
        }
        else
        {
            ConsoleLog.Log("Selected cleanup mode.");
        }

        var authentication = await OutlookAuthentication.Create(clientId);
        ConsoleLog.Log("Initialized Outlook authentication and persistent token-cache support.");

        if (login)
        {
            await authentication.Login();
            ConsoleLog.Log($"Login completed and the token cache was verified at {authentication.CacheFilePath}.");
            return 0;
        }

        var accessToken = await authentication.GetAccessToken();
        ConsoleLog.Log("Authentication completed without exposing the access token.");

        using var httpClient = new HttpClient();
        var cleaner = new JunkEmailService(httpClient, BlockedSenderNames.Values);
        var result = await cleaner.Clean(accessToken, dryRun);

        if (dryRun)
        {
            ConsoleLog.Log($"Dry run complete. {result.MatchedCount} of {result.ScannedCount} junk messages matched.");
        }
        else
        {
            ConsoleLog.Log($"Cleanup complete. Deleted {result.DeletedCount} of {result.ScannedCount} junk messages scanned.");
        }

        return 0;
    }

    private static void ShowHelp()
    {
        Console.WriteLine("""
            JunkEmailCleaner

            Usage:
              JunkEmailCleaner --login     Complete the one-time Outlook device login.
              JunkEmailCleaner --dry-run   Report matching junk messages without deleting them.
              JunkEmailCleaner             Delete matching messages from Junk Email.

            Required environment variable:
              JUNK_EMAIL_CLEANER_CLIENT_ID  Microsoft Entra application (client) ID.
            """);
    }
}
