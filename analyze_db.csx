#r "nuget: Microsoft.Data.Sqlite, 8.0.0"
using Microsoft.Data.Sqlite;

var dbPath = @"C:\Users\erik.richter\Desktop\MosaicTools\rvu_records.db";
using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
conn.Open();

// Get tables
Console.WriteLine("=== TABLES ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = "SELECT name FROM sqlite_master WHERE type='table'";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
        Console.WriteLine(reader.GetString(0));
}

// Get schema for each table
Console.WriteLine("\n=== SCHEMA ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = "SELECT sql FROM sqlite_master WHERE type='table'";
    using var reader = cmd.ExecuteReader();
    while (reader.Read())
        Console.WriteLine(reader.GetString(0) + "\n");
}

// Sample data from likely tables
foreach (var table in new[] { "shifts", "shift", "records", "rvu", "studies", "exams" })
{
    try
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT * FROM {table} LIMIT 5";
        using var reader = cmd.ExecuteReader();

        Console.WriteLine($"\n=== {table.ToUpper()} (sample) ===");

        // Column names
        for (int i = 0; i < reader.FieldCount; i++)
            Console.Write(reader.GetName(i) + "\t");
        Console.WriteLine();

        // Data
        while (reader.Read())
        {
            for (int i = 0; i < reader.FieldCount; i++)
                Console.Write(reader.GetValue(i)?.ToString() + "\t");
            Console.WriteLine();
        }
    }
    catch { }
}
