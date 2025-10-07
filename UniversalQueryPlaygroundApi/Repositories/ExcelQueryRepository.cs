using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
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

            var data = LoadSheet(workbook, request.Table);

            // Joins
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

            // Filter
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

            // Ordering
            if (!string.IsNullOrWhiteSpace(request.OrderBy))
            {
                var parts = request.OrderBy.Split(' ');
                var col = parts[0];
                var desc = parts.Length > 1 && parts[1].Equals("DESC", StringComparison.OrdinalIgnoreCase);

                Func<Dictionary<string, object>, object?> keySelector = row =>
                {
                    if (!row.TryGetValue(col, out var value) || value == null)
                        return null;

                    if (value is DateTime dt)
                        return dt;

                    if (double.TryParse(value.ToString(), out var dbl))
                        return dbl;

                    return value.ToString();
                };

                data = desc
                    ? data.OrderByDescending(keySelector, Comparer<object>.Create(CompareValues)).ToList()
                    : data.OrderBy(keySelector, Comparer<object>.Create(CompareValues)).ToList();
            }

            // Projection
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

            // ✅ Export results to a new sheet
            if (data.Any())
            {
                ExportResultToNewSheet(data);
            }

            return Task.FromResult<IEnumerable<Dictionary<string, object>>>(data);
        }

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
                            row.Cell(i + 1).GetString()))
                    .ToDictionary(kv => kv.Key, kv => kv.Value))
                .ToList();
        }

        private void ExportResultToNewSheet(List<Dictionary<string, object>> data)
        {
            using var workbook = new XLWorkbook(_excelPath);

            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var sheetName = $"QueryResult_{timestamp}";

            var ws = workbook.Worksheets.Add(sheetName);

            var headers = data.First().Keys.ToList();
            for (int i = 0; i < headers.Count; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
                ws.Cell(1, i + 1).Style.Font.Bold = true;
            }

            for (int r = 0; r < data.Count; r++)
            {
                int c = 0;
                foreach (var val in data[r].Values)
                {
                    ws.Cell(r + 2, c + 1).Value = val?.ToString() ?? string.Empty;
                    c++;
                }
            }

            var range = ws.Range(1, 1, data.Count + 1, headers.Count);
            var table = range.CreateTable();
            table.Theme = XLTableTheme.TableStyleMedium2;

            ws.Columns().AdjustToContents();

            workbook.Save();
        }

        private static int CompareValues(object? x, object? y)
        {
            if (x == null && y == null) return 0;
            if (x == null) return -1;
            if (y == null) return 1;

            if (x is IComparable cx && x.GetType() == y.GetType())
                return cx.CompareTo(y);

            return string.Compare(x.ToString(), y.ToString(), StringComparison.OrdinalIgnoreCase);
        }
    }
}
