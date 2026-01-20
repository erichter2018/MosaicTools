using Microsoft.Data.Sqlite;

var dbPath = @"C:\Users\erik.richter\Desktop\RVUCounter\csharp\RVUCounter\bin\Release\net8.0-windows\win-x64\publish\data\rvu_records.db";
Console.WriteLine($"Checking: {dbPath}");
Console.WriteLine($"Exists: {File.Exists(dbPath)}");

if (File.Exists(dbPath))
{
    // Don't use ReadOnly mode - it may not see WAL changes
    using var conn = new SqliteConnection($"Data Source={dbPath}");
    conn.Open();
    using var cmd = conn.CreateCommand();
    cmd.CommandText = "SELECT id, shift_start, is_current FROM shifts ORDER BY id DESC LIMIT 5";
    using var reader = cmd.ExecuteReader();
    Console.WriteLine("\nRecent shifts:");
    while (reader.Read())
    {
        Console.WriteLine($"  ID={reader.GetInt64(0)}, Start={reader.GetString(1)}, IsCurrent={reader.GetInt64(2)}");
    }
}
