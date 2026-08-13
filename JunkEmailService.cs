using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace JunkEmailCleaner;

internal sealed class JunkEmailService(HttpClient httpClient, IReadOnlyCollection<string> blockedSenderNames)
{
    private const string FirstPageUrl = "https://graph.microsoft.com/v1.0/me/mailFolders/junkemail/messages?$select=id,subject,from,sender,receivedDateTime&$top=100";
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    internal async Task<CleanupResult> Clean(string accessToken, bool dryRun)
    {
        ConsoleLog.Log(
            $"Starting Microsoft Graph junk-folder processing in {(dryRun ? "dry-run" : "delete")} mode with {blockedSenderNames.Count} configured sender-name filters.");

        var messages = await GetMessages(accessToken);
        ConsoleLog.Log($"Successfully downloaded metadata for {messages.Count} messages from the Junk Email folder.");

        var matches = messages.Where(IsBlockedSender).ToArray();
        ConsoleLog.Log($"Sender-name matching completed; {matches.Length} of {messages.Count} messages matched.");

        foreach (var message in matches)
        {
            var sender = GetSender(message);
            ConsoleLog.Log($"{(dryRun ? "Would delete" : "Deleting")}: {sender?.Name ?? "(no sender name)"} <{sender?.Address}> — {message.Subject ?? "(no subject)"}");

            if (!dryRun)
            {
                await DeleteMessage(accessToken, message.Id);
                ConsoleLog.Log($"Microsoft Graph confirmed deletion of the matched message from {sender?.Name ?? "(no sender name)"}.");
            }
        }

        if (matches.Length == 0)
        {
            ConsoleLog.Log("No messages matched, so no delete requests were necessary.");
        }

        return new CleanupResult(messages.Count, matches.Length, dryRun ? 0 : matches.Length);
    }

    private async Task<List<GraphMessage>> GetMessages(string accessToken)
    {
        var messages = new List<GraphMessage>();
        string? pageUrl = FirstPageUrl;
        var pageNumber = 0;

        while (pageUrl is not null)
        {
            pageNumber++;
            ConsoleLog.Log($"Requesting Junk Email message page {pageNumber} from Microsoft Graph.");
            using var request = CreateRequest(HttpMethod.Get, pageUrl, accessToken);
            using var response = await Send(request, $"download Junk Email message page {pageNumber}");
            await EnsureSuccess(response, $"download Junk Email message page {pageNumber}");
            ConsoleLog.Log(
                $"Microsoft Graph returned HTTP {(int)response.StatusCode} ({response.ReasonPhrase}) for message page {pageNumber}.");

            GraphMessagePage page;

            try
            {
                page = await response.Content.ReadFromJsonAsync<GraphMessagePage>(JsonOptions)
                    ?? throw new InvalidOperationException("Microsoft Graph returned an empty message page.");
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    $"Microsoft Graph returned a successful response for message page {pageNumber}, but its JSON could not be read: {exception.Message}",
                    exception);
            }

            messages.AddRange(page.Messages);
            ConsoleLog.Log(
                $"Parsed {page.Messages.Length} messages from page {pageNumber}; the running total is {messages.Count}.");
            pageUrl = page.NextLink;
        }

        ConsoleLog.Log($"Completed Microsoft Graph pagination after {pageNumber} page requests.");

        return messages;
    }

    private async Task DeleteMessage(string accessToken, string messageId)
    {
        var encodedMessageId = Uri.EscapeDataString(messageId);
        using var request = CreateRequest(HttpMethod.Delete, $"https://graph.microsoft.com/v1.0/me/messages/{encodedMessageId}", accessToken);
        using var response = await Send(request, "delete a matched Junk Email message");
        await EnsureSuccess(response, "delete a matched Junk Email message");

        if (response.StatusCode == HttpStatusCode.NoContent)
        {
            ConsoleLog.Log(
                "Microsoft Graph returned HTTP 204 (No Content). No Content means the email was deleted successfully.");
        }
        else
        {
            ConsoleLog.Log(
                $"Unexpected Response: Although Microsoft Graph returned the successful response HTTP {(int)response.StatusCode} ({response.ReasonPhrase}), the expected delete response was HTTP 204 (No Content).");
        }
    }

    private bool IsBlockedSender(GraphMessage message)
    {
        var senderName = GetSender(message)?.Name;

        return !string.IsNullOrWhiteSpace(senderName)
            && blockedSenderNames.Any(blockedName => senderName.Contains(blockedName, StringComparison.OrdinalIgnoreCase));
    }

    private static EmailAddress? GetSender(GraphMessage message) =>
        message.From?.EmailAddress ?? message.Sender?.EmailAddress;

    private static HttpRequestMessage CreateRequest(HttpMethod method, string url, string accessToken)
    {
        var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        return request;
    }

    private async Task<HttpResponseMessage> Send(HttpRequestMessage request, string operation)
    {
        try
        {
            return await httpClient.SendAsync(request);
        }
        catch (Exception exception)
        {
            throw new HttpRequestException(
                $"The request to {operation} failed before Microsoft Graph returned a response: {exception.Message}",
                exception);
        }
    }

    private static async Task EnsureSuccess(HttpResponseMessage response, string operation)
    {
        if (response.IsSuccessStatusCode)
        {
            return;
        }

        var responseBody = await response.Content.ReadAsStringAsync();
        throw new HttpRequestException(
            $"Microsoft Graph could not {operation} and returned HTTP {(int)response.StatusCode} ({response.ReasonPhrase}). Response body: {responseBody}",
            null,
            response.StatusCode);
    }
}

internal sealed record CleanupResult(int ScannedCount, int MatchedCount, int DeletedCount);

internal sealed record GraphMessagePage(
    [property: JsonPropertyName("value")] GraphMessage[] Messages,
    [property: JsonPropertyName("@odata.nextLink")] string? NextLink);

internal sealed record GraphMessage(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("subject")] string? Subject,
    [property: JsonPropertyName("from")] EmailAddressContainer? From,
    [property: JsonPropertyName("sender")] EmailAddressContainer? Sender,
    [property: JsonPropertyName("receivedDateTime")] DateTimeOffset? ReceivedDateTime);

internal sealed record EmailAddressContainer(
    [property: JsonPropertyName("emailAddress")] EmailAddress? EmailAddress);

internal sealed record EmailAddress(
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("address")] string? Address);
