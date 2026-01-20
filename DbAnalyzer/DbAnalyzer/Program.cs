using Microsoft.Data.Sqlite;

var dbPath = @"C:\Users\erik.richter\Desktop\MosaicTools\rvu_records.db";
using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
conn.Open();

// Find current shift
Console.WriteLine("=== CURRENT SHIFT ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = "SELECT id, shift_start, shift_end, is_current, shift_name FROM shifts WHERE is_current = 1";
    using var reader = cmd.ExecuteReader();
    if (reader.Read())
    {
        Console.WriteLine($"ID: {reader.GetInt64(0)}");
        Console.WriteLine($"Start: {reader.GetString(1)}");
        Console.WriteLine($"End: {(reader.IsDBNull(2) ? "NULL (ongoing)" : reader.GetString(2))}");
        Console.WriteLine($"Is Current: {reader.GetInt64(3)}");
        Console.WriteLine($"Name: {(reader.IsDBNull(4) ? "NULL" : reader.GetString(4))}");
    }
    else
    {
        Console.WriteLine("No current shift found (is_current = 1)");
    }
}

// Get RVU total for current shift
Console.WriteLine("\n=== CURRENT SHIFT RVU TOTAL ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = @"
        SELECT s.id, s.shift_start, COUNT(r.id) as record_count, COALESCE(SUM(r.rvu), 0) as total_rvu
        FROM shifts s
        LEFT JOIN records r ON r.shift_id = s.id
        WHERE s.is_current = 1
        GROUP BY s.id";
    using var reader = cmd.ExecuteReader();
    if (reader.Read())
    {
        Console.WriteLine($"Shift ID: {reader.GetInt64(0)}");
        Console.WriteLine($"Shift Start: {reader.GetString(1)}");
        Console.WriteLine($"Record Count: {reader.GetInt64(2)}");
        Console.WriteLine($"Total RVU: {reader.GetDouble(3):F2}");
    }
    else
    {
        Console.WriteLine("No current shift");
    }
}

// Also show the most recent shift (highest ID) as fallback
Console.WriteLine("\n=== MOST RECENT SHIFT (by ID) ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = @"
        SELECT s.id, s.shift_start, s.shift_end, s.is_current, COUNT(r.id) as record_count, COALESCE(SUM(r.rvu), 0) as total_rvu
        FROM shifts s
        LEFT JOIN records r ON r.shift_id = s.id
        GROUP BY s.id
        ORDER BY s.id DESC
        LIMIT 1";
    using var reader = cmd.ExecuteReader();
    if (reader.Read())
    {
        Console.WriteLine($"Shift ID: {reader.GetInt64(0)}");
        Console.WriteLine($"Shift Start: {reader.GetString(1)}");
        Console.WriteLine($"Shift End: {(reader.IsDBNull(2) ? "NULL (ongoing)" : reader.GetString(2))}");
        Console.WriteLine($"Is Current: {reader.GetInt64(3)}");
        Console.WriteLine($"Record Count: {reader.GetInt64(4)}");
        Console.WriteLine($"Total RVU: {reader.GetDouble(5):F2}");
    }
}

// Show a few sample RVU totals per shift
Console.WriteLine("\n=== SAMPLE SHIFT RVU TOTALS ===");
using (var cmd = conn.CreateCommand())
{
    cmd.CommandText = @"
        SELECT s.id, s.shift_start, s.is_current, COUNT(r.id) as record_count, COALESCE(SUM(r.rvu), 0) as total_rvu
        FROM shifts s
        LEFT JOIN records r ON r.shift_id = s.id
        GROUP BY s.id
        ORDER BY s.id DESC
        LIMIT 10";
    using var reader = cmd.ExecuteReader();
    Console.WriteLine("ID\tStart\t\t\t\tCurrent\tRecords\tRVU");
    Console.WriteLine(new string('-', 80));
    while (reader.Read())
    {
        Console.WriteLine($"{reader.GetInt64(0)}\t{reader.GetString(1)}\t{reader.GetInt64(2)}\t{reader.GetInt64(3)}\t{reader.GetDouble(4):F2}");
    }
}
