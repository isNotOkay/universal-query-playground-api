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

            // 1️⃣ Load base sheet
            var data = LoadSheet(workbook, request.Table);

            // 2️⃣ Apply joins (inner join by column)
            if (request.Joins != null)
            {
                foreach (var join in request.Joins)
                {
                    var right = LoadSheet(workbook, join.Table);
                    var leftKey = join.LeftColumn;
                    var rightKey = join.RightColumn;

                    data = (from l in data
                            join r in right
                            on l[leftKey] equals r[rightKey]
                            select l.Concat(r)
                                    // Handle duplicate columns: last one wins
                                    .GroupBy(kv => kv.Key)
                                    .Select(g => g.Last())
                                    .ToDictionary(kv => kv.Key, kv => kv.Value))
                           .ToList();
                }
            }

            // 3️⃣ Apply filtering (basic col = value)
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

            // 4️⃣ Apply ordering (before projection)
            if (!string.IsNullOrWhiteSpace(request.OrderBy))
            {
                var parts = request.OrderBy.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var col = parts[0];
                var desc = parts.Length > 1 && parts[1].Equals("DESC", StringComparison.OrdinalIgnoreCase);

                Func<Dictionary<string, object>, object?> keySelector = row =>
                {
                    if (!row.TryGetValue(col, out var value) || value == null)
                        return null;

                    if (DateTime.TryParse(value.ToString(), out var dt))
                        return dt;

                    if (double.TryParse(value.ToString(), out var dbl))
                        return dbl;

                    return value.ToString();
                };

                data = desc
                    ? data.OrderByDescending(keySelector, Comparer<object>.Create(CompareValues)).ToList()
                    : data.OrderBy(keySelector, Comparer<object>.Create(CompareValues)).ToList();
            }

            // 5️⃣ Projection (SELECT specific columns)
            if (request.Columns != null && request.Columns.Any())
            {
                data = data.Select(row =>
                    row.Where(kv => request.Columns.Contains(kv.Key))
                       .ToDictionary(kv => kv.Key, kv => kv.Value)
                ).ToList();
            }

            // 6️⃣ Offset & Limit (pagination)
            if (request.Offset.HasValue)
                data = data.Skip(request.Offset.Value).ToList();

            if (request.Limit.HasValue)
                data = data.Take(request.Limit.Value).ToList();

            // 7️⃣ Export to Excel sheet if requested
            if (data.Any() && !string.IsNullOrWhiteSpace(request.ExportSheetName))
            {
                ExportResultToSheet(data, request.ExportSheetName);
            }

            return Task.FromResult<IEnumerable<Dictionary<string, object>>>(data);
        }

        /// <summary>
        /// Loads a sheet into memory as a list of dictionaries (no column prefixes).
        /// </summary>
        private List<Dictionary<string, object>> LoadSheet(XLWorkbook workbook, string sheetName)
        {
            var sheet = workbook.Worksheet(sheetName);
            if (sheet == null)
                throw new ArgumentException($"Sheet '{sheetName}' not found in Excel workbook.");

            var headers = sheet.FirstRowUsed().Cells().Select(c => c.GetString()).ToList();

            return sheet.RangeUsed().RowsUsed().Skip(1)
                .Select(row => headers.Select((h, i) =>
                        new KeyValuePair<string, object>(
                            h, // ✅ no prefix anymore
                            row.Cell(i + 1).GetString()))
                    .ToDictionary(kv => kv.Key, kv => kv.Value))
                .ToList();
        }

        /// <summary>
        /// Exports query result to a new sheet with the given name. Replaces sheet if it exists.
        /// </summary>
        private void ExportResultToSheet(List<Dictionary<string, object>> data, string sheetName)
        {
            using var workbook = new XLWorkbook(_excelPath);

            // Replace existing sheet if present
            var existing = workbook.Worksheets.FirstOrDefault(s =>
                s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                workbook.Worksheets.Delete(existing.Name);
            }

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

        /// <summary>
        /// Safe comparer for mixed/null Excel values.
        /// </summary>
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
