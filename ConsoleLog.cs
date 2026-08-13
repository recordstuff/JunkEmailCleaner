using System.Globalization;

namespace JunkEmailCleaner;

internal static class ConsoleLog
{
    internal static void Log(string message) => Write(Console.Out, message);

    internal static void Error(string message) => Write(Console.Error, $"ERROR: {message}");

    internal static void Error(string message, Exception exception)
    {
        Error(message);
        Console.Error.WriteLine(exception);
    }

    private static void Write(TextWriter writer, string message)
    {
        var timestamp = DateTimeOffset.Now.ToString("G", CultureInfo.CurrentCulture);
        writer.WriteLine($"{timestamp} {message}");
    }
}
