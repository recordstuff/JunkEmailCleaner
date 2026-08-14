# Outlook Junk Email Cleaner

This project was recreated without history for security reasons.  However, this project didn't / does not contain any secrets.  There was some development history that is best moved somewhere else.

Outlook moves things to Junk before it runs rules.  My Junk folder is being spammed.  This utility gives me back my Junk folder.  With the bad emails gone, when a job offer or confirm your email or other important email goes to junk, now I can find it again.  Take that, spammer!  Phooey on spammers.

This is a one-shot .NET 8 console application that removes selected messages from the signed-in Outlook account's Junk Email folder. Sender display-name fragments are configured in `BlockedSenderNames.cs` and matched case-insensitively.

The app uses Microsoft Graph's message deletion API. Normal Outlook deletion and retention behavior applies. The app gathers all matches before deleting them so Graph pagination is not disrupted as the Junk Email folder changes.

## One-time Microsoft setup

1. Sign in to the [Azure portal](https://portal.azure.com/).
2. Make sure you are in a directory that you own, such as your personal directory. If you are not, use **Switch directory** to select one that you own.
3. Search for **App registrations**, select **New registration**, and name the application `JunkEmailCleaner`.
4. Choose **Personal Microsoft accounts only** as the supported account type.
5. Under **Authentication**, enable **Allow public client flows**.
6. Under **API permissions**, select **Add a permission** > **Microsoft Graph** > **Delegated permissions**, and add `Mail.ReadWrite`.
7. Copy the **Application (client) ID**. A client secret is neither required nor appropriate for this public console application.

Set the client ID in the environment:

```bash
export JUNK_EMAIL_CLEANER_CLIENT_ID="00000000-0000-0000-0000-000000000000"
```

## First login

From the project directory, run this interactively on the Linux account that should own the token cache:

```bash
dotnet run -- --login
```

The first `--` separates options for `dotnet run` from options passed to JunkEmailCleaner.

Follow the displayed device-login instructions. The delegated login is cached under the current user's local application-data directory. On Linux and macOS, the app restricts the directory to the current user and the cache file to read/write by the current user. The cache contains refresh credentials and must never be committed, copied casually, or shared.

When the cleaner and `--login` run as the same Linux account with a standard home directory, no cache configuration is needed. From the project directory, the simplest setup is:

```bash
export JUNK_EMAIL_CLEANER_CLIENT_ID="00000000-0000-0000-0000-000000000000"
dotnet run -- --login
dotnet run -- --dry-run
```

Set an explicit persistent cache directory when the server uses a nonstandard home environment or when the scheduled job runs as a different account from the account that created the cache:

```bash
export JUNK_EMAIL_CLEANER_CACHE_DIRECTORY="/home/example/.local/share/JunkEmailCleaner"
```

Replace `/home/example` with the actual home directory of the account that ran `--login`. Do not use `$HOME` for this case because it resolves to the home directory of the account currently running the cleaner, such as `/root` for a root cron job.

## Test and run

Each run writes timestamped progress messages. Timestamps use the computer's local timezone and regional date and time format. The log identifies every major configuration, authentication, cache, Microsoft Graph, matching, and deletion stage. Failures are marked with `ERROR:` and include the complete exception and inner-cause chain, but access tokens are never logged.

From the project directory, review matches without deleting anything:

```bash
dotnet run -- --dry-run
```

From the project directory, delete matching messages from Junk Email:

```bash
dotnet run
```

For cron, publish the application as a self-contained Linux executable and invoke the resulting executable. The machine that runs `dotnet publish` must have a compatible .NET SDK that can build the project's chosen target framework. This can be a development machine or the target Linux machine. After the self-contained executable has been published and copied to the target machine, that target does not need a separate .NET runtime installation to run it.

Ensure the cron environment provides both environment variables when an explicit cache directory is used.

```bash
dotnet publish JunkEmailCleaner.csproj -c Release -r linux-x64 --self-contained true -o ./publish
```

After publishing, `dotnet run --` is unnecessary. Run the published executable directly:

```bash
./publish/JunkEmailCleaner --login
./publish/JunkEmailCleaner --dry-run
./publish/JunkEmailCleaner
```

The following instructions are for the LINUX cron, but Windows Task Scheduler can be used instead on a Windows machine.

When cron runs as the same account that created the cache and that account has a standard home directory, the cron entry only needs the client ID:

```cron
*/15 * * * * JUNK_EMAIL_CLEANER_CLIENT_ID=00000000-0000-0000-0000-000000000000 /opt/junk-email-cleaner/JunkEmailCleaner >> /tmp/junk-email-cleaner.log 2>&1
```

When cron runs as `root` but `--login` created the cache for a normal user, specify that user's cache directory explicitly. Without this setting, the cleaner looks under root's home directory instead. Root can read the normal user's owner-restricted cache, and setting the cache directory does not change its ownership.

The same explicit setting applies to a nonstandard home environment:

```cron
*/15 * * * * JUNK_EMAIL_CLEANER_CLIENT_ID=00000000-0000-0000-0000-000000000000 JUNK_EMAIL_CLEANER_CACHE_DIRECTORY=/home/example/.local/share/JunkEmailCleaner /opt/junk-email-cleaner/JunkEmailCleaner >> /var/log/junk-email-cleaner.log 2>&1
```

If Microsoft revokes the cached consent or refresh token, the unattended run exits with code `3`; rerun with `--login` interactively.
