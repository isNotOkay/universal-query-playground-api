using UniversalQueryPlaygroundApi.Models;

namespace UniversalQueryPlaygroundApi.Repositories
{
    public interface IQueryRepository
    {
        Task<IEnumerable<Dictionary<string, object>>> ExecuteAsync(QueryRequest request);
    }
}