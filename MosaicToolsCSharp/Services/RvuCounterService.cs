using Microsoft.Data.Sqlite;

namespace MosaicTools.Services;

/// <summary>
/// Reads RVU data from a local RVUCounter SQLite database.
/// </summary>
public class RvuCounterService
{
    private readonly Configuration _config;

    public RvuCounterService(Configuration config)
    {
        _config = config;
    }

    /// <summary>
    /// Get the database path from config (now stores the full path directly).
    /// </summary>
    private string? GetDatabasePath()
    {
        if (string.IsNullOrEmpty(_config.RvuCounterPath))
            return null;

        return File.Exists(_config.RvuCounterPath) ? _config.RvuCounterPath : null;
    }

    /// <summary>
    /// Check if the RVUCounter database is accessible.
    /// </summary>
    public bool IsAvailable()
    {
        return GetDatabasePath() != null;
    }

    /// <summary>
    /// Get the current shift RVU total.
    /// Returns null if no current shift or database not available.
    /// </summary>
    public double? GetCurrentShiftRvuTotal()
    {
        var dbPath = GetDatabasePath();
        if (dbPath == null)
        {
            Logger.Trace("RvuCounterService: Database not found");
            return null;
        }

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            using var cmd = conn.CreateCommand();

            // First check if there's a current shift (pick most recent if multiple)
            cmd.CommandText = "SELECT id FROM shifts WHERE is_current = 1 ORDER BY id DESC LIMIT 1";
            Logger.Trace($"RvuCounterService: Querying database at {dbPath}");
            var shiftIdResult = cmd.ExecuteScalar();
            Logger.Trace($"RvuCounterService: Shift query result = {shiftIdResult ?? "NULL"}");
            if (shiftIdResult == null || shiftIdResult == DBNull.Value)
            {
                Logger.Trace("RvuCounterService: No current shift found");
                return null;
            }

            // Get the RVU total for that specific shift
            cmd.CommandText = "SELECT COALESCE(SUM(rvu), 0) FROM records WHERE shift_id = @shiftId";
            cmd.Parameters.AddWithValue("@shiftId", shiftIdResult);

            var result = cmd.ExecuteScalar();
            var total = Convert.ToDouble(result ?? 0);
            Logger.Trace($"RvuCounterService: Current shift RVU total = {total:F2}");
            return total;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RvuCounterService: Error reading database: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Get current shift info including RVU total and record count.
    /// </summary>
    public ShiftInfo? GetCurrentShiftInfo()
    {
        var dbPath = GetDatabasePath();
        if (dbPath == null)
            return null;

        try
        {
            using var conn = new SqliteConnection($"Data Source={dbPath};Mode=ReadOnly");
            conn.Open();

            using var cmd = conn.CreateCommand();

            // First check if there's a current shift (pick most recent if multiple)
            cmd.CommandText = "SELECT id FROM shifts WHERE is_current = 1 ORDER BY id DESC LIMIT 1";
            var shiftIdResult = cmd.ExecuteScalar();
            if (shiftIdResult == null || shiftIdResult == DBNull.Value)
            {
                return null;
            }

            // Get shift info for that specific shift
            cmd.CommandText = @"
                SELECT s.id, s.shift_start, COUNT(r.id) as record_count, COALESCE(SUM(r.rvu), 0) as total_rvu
                FROM shifts s
                LEFT JOIN records r ON r.shift_id = s.id
                WHERE s.id = @shiftId
                GROUP BY s.id";
            cmd.Parameters.AddWithValue("@shiftId", shiftIdResult);

            using var reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                return new ShiftInfo
                {
                    ShiftId = reader.GetInt64(0),
                    ShiftStart = reader.GetString(1),
                    RecordCount = (int)reader.GetInt64(2),
                    TotalRvu = reader.GetDouble(3)
                };
            }

            return null;
        }
        catch (Exception ex)
        {
            Logger.Trace($"RvuCounterService: Error reading shift info: {ex.Message}");
            return null;
        }
    }
}

/// <summary>
/// Information about a shift from RVUCounter.
/// </summary>
public class ShiftInfo
{
    public long ShiftId { get; set; }
    public string ShiftStart { get; set; } = "";
    public int RecordCount { get; set; }
    public double TotalRvu { get; set; }
}
