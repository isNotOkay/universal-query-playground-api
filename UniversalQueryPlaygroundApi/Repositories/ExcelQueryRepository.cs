using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using UniversalQueryPlaygroundApi.Models;

namespace UniversalQueryPlaygroundApi.Repositories
{
    public class ExcelQueryRepository : IQueryRepository
    {
        private readonly string _excelPath;

        public ExcelQueryRepository(IConfiguration config)
        {
            _excelPath = config["DataSources:ExcelFile"] 
                ?? throw new Exception("Excel file path not configured (DataSources:ExcelFile).");

            if (!File.Exists(_excelPath))
                throw new FileNotFoundException($"Excel file not found: {_excelPath}");
        }

        public Task<IEnumerable<Dictionary<string, object>>> ExecuteAsync(QueryRequest request)
        {
            using var workbook = new XLWorkbook(_excelPath);

            // Load base table (sheet)
            var data = LoadSheet(workbook, request.Table);

            // Apply JOINs (always INNER JOIN by column)
            if (request.Joins != null)
            {
                foreach (var join in request.Joins)
                {
                    var right = LoadSheet(workbook, join.Table);
                    var leftKey = $"{request.Table}.{join.LeftColumn}";
                    var rightKey = $"{join.Table}.{join.RightColumn}";

                    data = (from l in data
                            join r in right
                            on l[leftKey] equals r[rightKey]
                            select l.Concat(r)
                                    .ToDictionary(kv => kv.Key, kv => kv.Value))
                           .ToList();
                }
            }

            // Apply filter (simple col = value for now)
            if (!string.IsNullOrWhiteSpace(request.Filter))
            {
                var parts = request.Filter.Split('=', 2);
                if (parts.Length == 2)
                {
                    var col = parts[0].Trim();
                    var val = parts[1].Trim().Trim('\'', '"');

                    data = data.Where(r =>
                        r.ContainsKey(col) &&
                        r[col]?.ToString()?.Equals(val, StringComparison.OrdinalIgnoreCase) == true
                    ).ToList();
                }
            }

            // Apply ordering BEFORE projection for robustness
            if (!string.IsNullOrWhiteSpace(request.OrderBy))
            {
                var parts = request.OrderBy.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var col = parts[0];
                var desc = parts.Length > 1 && parts[1].Equals("DESC", StringComparison.OrdinalIgnoreCase);

                Func<Dictionary<string, object>, object?> keySelector = row =>
                {
                    if (!row.TryGetValue(col, out var value) || value == null)
                        return null;

                    // Normalize common types for comparison
                    if (value is DateTime dt)
                        return dt;

                    if (double.TryParse(value.ToString(), out var dbl))
                        return dbl;

                    return value.ToString(); // fallback to string
                };

                data = desc
                    ? data.OrderByDescending(keySelector, Comparer<object>.Create(CompareValues)).ToList()
                    : data.OrderBy(keySelector, Comparer<object>.Create(CompareValues)).ToList();
            }

            // Projection (SELECT specific columns)
            if (request.Columns != null && request.Columns.Any())
            {
                data = data.Select(row =>
                    row.Where(kv => request.Columns.Contains(kv.Key))
                       .ToDictionary(kv => kv.Key, kv => kv.Value)
                ).ToList();
            }

            // Offset & Limit
            if (request.Offset.HasValue)
                data = data.Skip(request.Offset.Value).ToList();

            if (request.Limit.HasValue)
                data = data.Take(request.Limit.Value).ToList();

            return Task.FromResult<IEnumerable<Dictionary<string, object>>>(data);
        }

        // Helper: Load one Excel sheet into memory with namespaced columns
        private List<Dictionary<string, object>> LoadSheet(XLWorkbook workbook, string sheetName)
        {
            var sheet = workbook.Worksheet(sheetName);
            if (sheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found in Excel workbook.");

            var headers = sheet.FirstRowUsed().Cells().Select(c => c.GetString()).ToList();

            return sheet.RangeUsed().RowsUsed().Skip(1)
                .Select(row => headers.Select((h, i) =>
                        new KeyValuePair<string, object>(
                            $"{sheetName}.{h}",
                            row.Cell(i + 1).GetString() // always returns string, never null
                        ))
                    .ToDictionary(kv => kv.Key, kv => kv.Value))
                .ToList();
        }

        // Helper: Safe comparer for mixed/null Excel values
        private static int CompareValues(object? x, object? y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            // Compare same types directly
            if (x is IComparable cx && x.GetType() == y.GetType())
                return cx.CompareTo(y);

            // Fallback: compare as strings
            return string.Compare(x.ToString(), y.ToString(), StringComparison.OrdinalIgnoreCase);
        }
    }
}
