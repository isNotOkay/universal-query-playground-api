using UniversalQueryPlaygroundApi.Models;
using UniversalQueryPlaygroundApi.Repositories;

namespace UniversalQueryPlaygroundApi.Services
{
    public class QueryService(SqliteQueryRepository sqliteRepo, ExcelQueryRepository excelRepo)
    {
        public async Task<IEnumerable<Dictionary<string, object>>> ExecuteAsync(QueryRequest req)
        {
            return req.Engine.ToLowerInvariant() switch
            {
                "sqlite" => await sqliteRepo.ExecuteAsync(req),
                "excel" => await excelRepo.ExecuteAsync(req),
                _ => throw new NotSupportedException($"Engine {req.Engine} is not supported.")
            };
        }
    }
}