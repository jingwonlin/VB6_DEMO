using Microsoft.Data.SqlClient;
using System.Data;

namespace GF3Com.Database;

/// <summary>
/// 資料庫存取層 (replaces ADODB.Connection / Recordset in Global.bas and SK900_plc.bas)
/// </summary>
internal sealed class DatabaseContext : IDisposable
{
    private SqlConnection? _rcn;
    private SqlConnection? _rcn1;
    private bool _disposed;

    public bool IsOpen => _rcn?.State == ConnectionState.Open;

    /// <summary>
    /// 開啟主連線與次要連線 (rcn, rcn1 in VB6)
    /// </summary>
    public void Open(string connectionString)
    {
        _rcn  = new SqlConnection(connectionString);
        _rcn1 = new SqlConnection(connectionString);

        _rcn.Open();
        _rcn1.Open();
    }

    /// <summary>執行不回傳資料的 SQL（UPDATE / INSERT / DELETE）</summary>
    public void Execute(string sql, SqlConnection? conn = null)
    {
        var target = conn ?? _rcn ?? throw new InvalidOperationException("DB not open");
        using var cmd = new SqlCommand(sql, target);
        cmd.ExecuteNonQuery();
    }

    /// <summary>以次要連線執行（rcn1 in VB6）</summary>
    public void Execute1(string sql) => Execute(sql, _rcn1);

    /// <summary>查詢並回傳 DataTable</summary>
    public DataTable Query(string sql, SqlConnection? conn = null)
    {
        var target = conn ?? _rcn ?? throw new InvalidOperationException("DB not open");
        using var adapter = new SqlDataAdapter(sql, target);
        var dt = new DataTable();
        adapter.Fill(dt);
        return dt;
    }

    /// <summary>查詢單一純量值</summary>
    public object? QueryScalar(string sql, SqlConnection? conn = null)
    {
        var target = conn ?? _rcn ?? throw new InvalidOperationException("DB not open");
        using var cmd = new SqlCommand(sql, target);
        return cmd.ExecuteScalar();
    }

    /// <summary>以參數化查詢防止 SQL Injection</summary>
    public void ExecuteParameterized(string sql,
        IEnumerable<(string name, object? value)> parameters,
        SqlConnection? conn = null)
    {
        var target = conn ?? _rcn ?? throw new InvalidOperationException("DB not open");
        using var cmd = new SqlCommand(sql, target);
        foreach (var (name, value) in parameters)
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _rcn?.Dispose();
        _rcn1?.Dispose();
        _disposed = true;
    }
}
