using System.Text;
using Microsoft.Data.Sqlite;
using UniversalQueryPlaygroundApi.Models;

namespace UniversalQueryPlaygroundApi.Repositories
{
    public class SqliteQueryRepository(IConfiguration config) : IQueryRepository
    {
        private readonly string _dbPath = config["DataSources:SqliteFile"] ?? throw new Exception("Sqlite file path not configured.");

        public async Task<IEnumerable<Dictionary<string, object>>> ExecuteAsync(QueryRequest request)
        {
            var sql = BuildSql(request);

            var results = new List<Dictionary<string, object>>();
            var connString = $"Data Source={_dbPath}";

            await using var conn = new SqliteConnection(connString);
            await conn.OpenAsync();
            await using var cmd = new SqliteCommand(sql, conn);
            await using var reader = await cmd.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[reader.GetName(i)] = reader.GetValue(i);
                }
                results.Add(row);
            }

            return results;
        }

        private string BuildSql(QueryRequest req)
        {
            var sb = new StringBuilder();
            sb.Append("SELECT ");

            sb.Append(req.Columns != null && req.Columns.Any()
                ? string.Join(", ", req.Columns)
                : "*");

            sb.Append($" FROM {req.Table}");

            if (req.Joins != null)
            {
                foreach (var j in req.Joins)
                {
                    sb.Append($" INNER JOIN {j.Table} ON {req.Table}.{j.LeftColumn} = {j.Table}.{j.RightColumn}");
                }
            }

            if (!string.IsNullOrWhiteSpace(req.Filter))
                sb.Append($" WHERE {req.Filter}");

            if (!string.IsNullOrWhiteSpace(req.OrderBy))
                sb.Append($" ORDER BY {req.OrderBy}");

            if (req.Limit.HasValue)
                sb.Append($" LIMIT {req.Limit.Value}");

            if (req.Offset.HasValue)
                sb.Append($" OFFSET {req.Offset.Value}");

            return sb.ToString();
        }
    }
}
