using Microsoft.AspNetCore.Mvc;
using UniversalQueryPlaygroundApi.Models;
using UniversalQueryPlaygroundApi.Services;

namespace UniversalQueryPlaygroundApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class QueryController(QueryService service, ILogger<QueryController> logger) : ControllerBase
    {
        private readonly ILogger<QueryController> _logger = logger;

        [HttpPost]
        public async Task<IActionResult> Post([FromBody] QueryRequest request)
        {
            var result = await service.ExecuteAsync(request);
            return Ok(result);
        }
    }
}