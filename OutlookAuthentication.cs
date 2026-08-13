using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace JunkEmailCleaner;

internal sealed class OutlookAuthentication
{
    private const string CacheFileName = "msal-cache.bin";
    private static readonly string[] Scopes = ["Mail.ReadWrite"];

    private readonly IPublicClientApplication application;
    private readonly MsalCacheHelper cacheHelper;
    private readonly string cacheDirectory;
    private readonly string clientId;

    private OutlookAuthentication(
        IPublicClientApplication application,
        MsalCacheHelper cacheHelper,
        string cacheDirectory,
        string clientId)
    {
        this.application = application;
        this.cacheHelper = cacheHelper;
        this.cacheDirectory = cacheDirectory;
        this.clientId = clientId;
    }

    internal string CacheFilePath => Path.Combine(cacheDirectory, CacheFileName);

    internal static async Task<OutlookAuthentication> Create(string clientId)
    {
        var cacheDirectory = GetCacheDirectory();
        ConsoleLog.Log($"Resolved the Outlook token-cache directory to {cacheDirectory}.");

        try
        {
            Directory.CreateDirectory(cacheDirectory);
            ConsoleLog.Log($"Verified that the token-cache directory exists at {cacheDirectory}.");
            RestrictDirectoryPermissions(cacheDirectory);
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"Token-cache directory setup failed at {cacheDirectory}: {exception.Message}",
                exception);
        }

        var application = PublicClientApplicationBuilder
            .Create(clientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, "consumers")
            .WithLegacyCacheCompatibility(false)
            .Build();
        ConsoleLog.Log("Created the Microsoft public-client application for personal Microsoft accounts.");

        var storageProperties = new StorageCreationPropertiesBuilder(CacheFileName, cacheDirectory)
            .WithUnprotectedFile()
            .Build();
        ConsoleLog.Log(
            $"Configured an unencrypted token-cache file at {Path.Combine(cacheDirectory, CacheFileName)}; operating-system permissions restrict access.");

        MsalCacheHelper cacheHelper;

        try
        {
            cacheHelper = await MsalCacheHelper.CreateAsync(storageProperties);
            cacheHelper.RegisterCache(application.UserTokenCache);
            ConsoleLog.Log("Registered Microsoft MSAL token-cache persistence.");
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"MSAL token-cache initialization failed at {cacheDirectory}: {exception.Message}",
                exception);
        }

        return new OutlookAuthentication(application, cacheHelper, cacheDirectory, clientId);
    }

    internal async Task Login()
    {
        ConsoleLog.Log($"Starting interactive Outlook login. The token cache will be saved to {CacheFilePath}.");

        var result = await application
            .AcquireTokenWithDeviceCode(Scopes, deviceCode =>
            {
                ConsoleLog.Log("Microsoft supplied a device-login code and is waiting for browser confirmation.");
                Console.WriteLine(deviceCode.Message);
                return Task.CompletedTask;
            })
            .ExecuteAsync();
        ConsoleLog.Log(
            $"Microsoft completed the interactive login for account {result.Account.Username}; the returned access token was not logged.");

        RestrictCacheFilePermissions();

        if (!File.Exists(CacheFilePath))
        {
            throw new InvalidOperationException(
                $"Microsoft completed the Outlook login, but MSAL did not create the expected token-cache file at {CacheFilePath}.");
        }

        var cacheLength = new FileInfo(CacheFilePath).Length;

        if (cacheLength == 0)
        {
            throw new InvalidOperationException(
                $"Microsoft completed the Outlook login, but the token-cache file at {CacheFilePath} is empty.");
        }

        ConsoleLog.Log($"Verified the token-cache file at {CacheFilePath}; it contains {cacheLength} bytes.");

        var accountCount = (await application.GetAccountsAsync()).Count();

        if (accountCount == 0)
        {
            throw new InvalidOperationException(
                $"Microsoft completed the Outlook login and wrote {cacheLength} bytes to {CacheFilePath}, but MSAL could not reload an account from that cache.");
        }

        ConsoleLog.Log(
            $"Verified that MSAL can reload {accountCount} {(accountCount == 1 ? "account" : "accounts")} from the saved token cache.");
    }

    internal async Task<string> GetAccessToken()
    {
        ConsoleLog.Log($"Looking for the Outlook token cache at {CacheFilePath}.");

        var cacheFile = new FileInfo(CacheFilePath);

        if (!cacheFile.Exists)
        {
            throw new InvalidOperationException(
                $"Could not find the Outlook token cache at {CacheFilePath}. Run --login as the user that will run the cleaner.");
        }

        ConsoleLog.Log($"Found the Outlook token cache at {CacheFilePath}.");

        byte[] cacheData;

        try
        {
            cacheData = cacheHelper.LoadUnencryptedTokenCache();
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"The Outlook token cache at {CacheFilePath} could not be read by user {Environment.UserName}: {exception.Message}",
                exception);
        }

        if (cacheData.Length == 0)
        {
            throw new InvalidOperationException($"The Outlook token cache at {CacheFilePath} is empty.");
        }

        ConsoleLog.Log(
            $"Successfully read {cacheData.Length} bytes from the Outlook token cache at {CacheFilePath} as user {Environment.UserName}.");

        IAccount[] accounts;

        try
        {
            accounts = (await application.GetAccountsAsync()).ToArray();
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"MSAL read {cacheData.Length} bytes from {CacheFilePath}, but failed while deserializing cached accounts: {exception.Message}",
                exception);
        }

        if (accounts.Length == 0)
        {
            throw new InvalidOperationException(
                $"The Outlook token cache at {CacheFilePath} contains {cacheFile.Length} bytes, but it has no account for client ID {clientId}. Use the same JUNK_EMAIL_CLEANER_CLIENT_ID for --login and scheduled runs.");
        }

        if (accounts.Length > 1)
        {
            throw new InvalidOperationException("More than one Outlook account exists in the token cache. Remove the cache and run --login again.");
        }

        ConsoleLog.Log("Successfully loaded one Outlook account from the token cache.");

        AuthenticationResult result;

        try
        {
            result = await application
                .AcquireTokenSilent(Scopes, accounts[0])
                .ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new InvalidOperationException(
                $"MSAL loaded the cached account but failed to acquire Outlook access silently: {exception.Message}",
                exception);
        }

        RestrictCacheFilePermissions();
        ConsoleLog.Log(
            $"Successfully acquired Outlook access using the saved login; token source was {result.AuthenticationResultMetadata.TokenSource}.");
        return result.AccessToken;
    }

    private static string GetCacheDirectory()
    {
        var configuredDirectory = Environment.GetEnvironmentVariable("JUNK_EMAIL_CLEANER_CACHE_DIRECTORY");

        if (!string.IsNullOrWhiteSpace(configuredDirectory))
        {
            var fullPath = Path.GetFullPath(configuredDirectory);
            ConsoleLog.Log(
                $"Found JUNK_EMAIL_CLEANER_CACHE_DIRECTORY and resolved its value to {fullPath}.");
            return fullPath;
        }

        ConsoleLog.Log("JUNK_EMAIL_CLEANER_CACHE_DIRECTORY is not set; resolving the current user's standard local application-data directory.");

        var localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        if (string.IsNullOrWhiteSpace(localApplicationData))
        {
            var homeDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            if (string.IsNullOrWhiteSpace(homeDirectory))
            {
                throw new InvalidOperationException(
                    "The token cache directory could not be determined. Set JUNK_EMAIL_CLEANER_CACHE_DIRECTORY explicitly.");
            }

            localApplicationData = Path.Combine(homeDirectory, ".local", "share");
            ConsoleLog.Log(
                $"LocalApplicationData was empty, so the token-cache base directory was derived from the user profile as {localApplicationData}.");
        }
        else
        {
            ConsoleLog.Log($"Resolved the standard local application-data directory to {localApplicationData}.");
        }

        return Path.Combine(localApplicationData, "JunkEmailCleaner");
    }

    private static void RestrictDirectoryPermissions(string directory)
    {
        if (OperatingSystem.IsLinux() || OperatingSystem.IsMacOS())
        {
            File.SetUnixFileMode(directory, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            ConsoleLog.Log($"Restricted token-cache directory permissions to the owning user at {directory}.");
        }
        else
        {
            ConsoleLog.Log("Skipped Unix token-cache directory permissions because this operating system is not Linux or macOS.");
        }
    }

    private void RestrictCacheFilePermissions()
    {
        if (File.Exists(CacheFilePath) && (OperatingSystem.IsLinux() || OperatingSystem.IsMacOS()))
        {
            File.SetUnixFileMode(CacheFilePath, UnixFileMode.UserRead | UnixFileMode.UserWrite);
            ConsoleLog.Log($"Restricted token-cache file permissions to the owning user at {CacheFilePath}.");
        }
        else if (!File.Exists(CacheFilePath))
        {
            ConsoleLog.Log($"Token-cache file permissions were not changed because no file exists yet at {CacheFilePath}.");
        }
        else
        {
            ConsoleLog.Log("Skipped Unix token-cache file permissions because this operating system is not Linux or macOS.");
        }
    }
}
