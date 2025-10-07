namespace UniversalQueryPlaygroundApi.Models
{
    public class JoinRequest
    {
        public string Table { get; set; } = string.Empty;
        public string LeftColumn { get; set; } = string.Empty;
        public string RightColumn { get; set; } = string.Empty;
    }

    public class QueryRequest
    {
        public string Engine { get; set; } = string.Empty; // "sqlite" or "excel"
        public string Table { get; set; } = string.Empty;
        public List<string>? Columns { get; set; }
        public List<JoinRequest>? Joins { get; set; }
        public string? Filter { get; set; }
        public string? OrderBy { get; set; }
        public int? Limit { get; set; }
        public int? Offset { get; set; }
    }
}