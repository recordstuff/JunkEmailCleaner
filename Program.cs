using JunkEmailCleaner;

try
{
    return await Application.Run(args);
}
catch (Exception exception)
{
    try
    {
        ConsoleLog.Error(
            "JunkEmailCleaner reached its final safety handler after an exception escaped Application.Run.",
            exception);
    }
    catch
    {
        // The primary logging destination is itself unavailable. Return a failure code without throwing again.
    }

    return 1;
}
