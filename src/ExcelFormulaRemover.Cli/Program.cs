namespace ExcelFormulaRemover.Cli;

using System.Runtime.InteropServices;
using System.Text;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        using CancellationTokenSource cancellation = new();
        int cancelSignalReceived = 0;
        ConsoleCancelEventHandler cancelHandler = (_, eventArgs) =>
        {
            if (Interlocked.Exchange(ref cancelSignalReceived, 1) == 0)
            {
                eventArgs.Cancel = true;
                cancellation.Cancel();
            }
            else
            {
                eventArgs.Cancel = false;
            }
        };

        // Harnesses and `timeout` stop a process with SIGTERM rather than
        // Ctrl+C; treat it as the same cancellation so the result for files
        // already completed is still written.
        using PosixSignalRegistration termination = PosixSignalRegistration.Create(
            PosixSignal.SIGTERM,
            context =>
            {
                context.Cancel = true;
                cancellation.Cancel();
            });

        Console.CancelKeyPress += cancelHandler;
        try
        {
            return await CliApplication.RunAsync(
                args,
                Console.Out,
                Console.Error,
                cancellation.Token).ConfigureAwait(false);
        }
        finally
        {
            Console.CancelKeyPress -= cancelHandler;
        }
    }
}
