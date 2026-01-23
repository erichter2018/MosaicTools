using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ClarioIgnore;

static class Program
{
    private static Mutex? _mutex;
    private const string MutexName = "ClarioIgnore_SingleInstance_Mutex";

    [DllImport("kernel32.dll")]
    private static extern bool AllocConsole();

    [DllImport("kernel32.dll")]
    private static extern bool FreeConsole();

    [STAThread]
    static void Main(string[] args)
    {
        // Debug mode - run once and show results in console
        if (args.Length > 0 && args[0] == "-debug")
        {
            AllocConsole();
            try
            {
                RunDebugMode();
            }
            finally
            {
                FreeConsole();
            }
            return;
        }

        // Single instance check
        _mutex = new Mutex(true, MutexName, out bool createdNew);
        if (!createdNew)
        {
            MessageBox.Show("ClarioIgnore is already running.", "ClarioIgnore",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        try
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
        finally
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
        }
    }

    private static void RunDebugMode()
    {
        Console.WriteLine("ClarioIgnore Debug Mode");
        Console.WriteLine("========================\n");

        using var service = new ClarioService();

        Console.WriteLine("Looking for Clario window...");
        var window = service.FindClarioWindow();

        if (window == null)
        {
            Console.WriteLine("ERROR: Clario window not found.");
            Console.WriteLine("Make sure Chrome is open with Clario worklist.");
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
            return;
        }

        Console.WriteLine($"Found Clario window: '{window.Name}'\n");

        Console.WriteLine("Reading worklist items...\n");
        var items = service.GetWorklistItems();

        if (items.Count == 0)
        {
            Console.WriteLine("No worklist items found.");
        }
        else
        {
            Console.WriteLine($"Found {items.Count} worklist item(s):\n");

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                Console.WriteLine($"[{i + 1}] Procedure: {item.Procedure}");
                if (!string.IsNullOrEmpty(item.Priority))
                    Console.WriteLine($"    Priority: {item.Priority}");
                if (!string.IsNullOrEmpty(item.Accession))
                    Console.WriteLine($"    Accession: {item.Accession}");

                // Check against rules
                var rule = Configuration.Instance.FindMatchingRule(item.Procedure);
                if (rule != null)
                    Console.WriteLine($"    *** MATCHES RULE: {rule.Name} ***");

                Console.WriteLine();
            }
        }

        Console.WriteLine("\n--- Skip Rules Configured ---");
        foreach (var rule in Configuration.Instance.SkipRules)
        {
            var status = rule.Enabled ? "ENABLED" : "disabled";
            Console.WriteLine($"[{status}] {rule.Name}");
            if (!string.IsNullOrEmpty(rule.CriteriaRequired))
                Console.WriteLine($"  Required: {rule.CriteriaRequired}");
            if (!string.IsNullOrEmpty(rule.CriteriaAnyOf))
                Console.WriteLine($"  Any Of: {rule.CriteriaAnyOf}");
            if (!string.IsNullOrEmpty(rule.CriteriaExclude))
                Console.WriteLine($"  Exclude: {rule.CriteriaExclude}");
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}
