using UniversalQueryPlaygroundApi.Repositories;
using UniversalQueryPlaygroundApi.Services;

var builder = WebApplication.CreateBuilder(args);

// Add controllers
builder.Services.AddControllers();

// Add OpenAPI support (built-in since .NET 8+)
builder.Services.AddOpenApi();

// Register application services
builder.Services.AddScoped<SqliteQueryRepository>();
builder.Services.AddScoped<ExcelQueryRepository>();
builder.Services.AddScoped<QueryService>();

var app = builder.Build();

// Enable OpenAPI UI in development
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi(); // serves OpenAPI at /openapi/v1.json by default
    app.UseDeveloperExceptionPage();
}

app.UseStaticFiles();

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();