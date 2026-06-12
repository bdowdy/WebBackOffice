using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers;

[ApiController]
[Route("api/discounts")]
public sealed class DiscountsController : ControllerBase
{
    private const string SelectDiscountsSql = """
        SELECT
            [Discount Name],
            [Discount Type],
            [Discount Amount],
            [Access Level],
            [Category Link],
            [Menu Item Link],
            [Max Discount Amount],
            [Match],
            [Taxable],
            [Allow Only One],
            [After Print],
            [Department Link],
            [Auto Match],
            [Disc Department],
            [Type],
            [Search Number],
            [Start Time],
            [End Time],
            [Auto Deduct],
            [Tax 1],
            [Tax 2],
            [Expires],
            [Discount Id],
            [Tracking]
        FROM [dbo].[Discounts]
        ORDER BY [Discount Name];
        """;

    private readonly IConfiguration _configuration;
    private readonly ILogger<DiscountsController> _logger;

    public DiscountsController(
        IConfiguration configuration,
        ILogger<DiscountsController> logger)
    {
        _configuration = configuration;
        _logger = logger;
    }

    [HttpGet]
    public async Task<ActionResult<IReadOnlyList<DiscountDto>>> GetDiscounts(CancellationToken cancellationToken)
    {
        var connectionString = _configuration.GetConnectionString("DefaultConnection");

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            _logger.LogError("DefaultConnection connection string is not configured.");
            return StatusCode(StatusCodes.Status500InternalServerError, "Server configuration error.");
        }

        try
        {
            await using var connection = new SqlConnection(connectionString);
            await connection.OpenAsync(cancellationToken);

            await using var command = new SqlCommand(SelectDiscountsSql, connection);
            await using var reader = await command.ExecuteReaderAsync(cancellationToken);

            var discounts = new List<DiscountDto>();
            var columns = new DiscountColumns(reader);

            while (await reader.ReadAsync(cancellationToken))
            {
                discounts.Add(columns.MapRow(reader));
            }

            return Ok(discounts);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to load discounts.");
            return StatusCode(StatusCodes.Status500InternalServerError, "Unable to load discounts.");
        }
    }

    /// <summary>
    /// Caches column ordinals once per reader so we avoid the per-row string lookups
    /// the previous implementation paid 24 times per row.
    /// </summary>
    private sealed class DiscountColumns
    {
        private readonly int _discountName;
        private readonly int _discountType;
        private readonly int _discountAmount;
        private readonly int _accessLevel;
        private readonly int _categoryLink;
        private readonly int _menuItemLink;
        private readonly int _maxDiscountAmount;
        private readonly int _match;
        private readonly int _taxable;
        private readonly int _allowOnlyOne;
        private readonly int _afterPrint;
        private readonly int _departmentLink;
        private readonly int _autoMatch;
        private readonly int _discDepartment;
        private readonly int _type;
        private readonly int _searchNumber;
        private readonly int _startTime;
        private readonly int _endTime;
        private readonly int _autoDeduct;
        private readonly int _tax1;
        private readonly int _tax2;
        private readonly int _expires;
        private readonly int _discountId;
        private readonly int _tracking;

        public DiscountColumns(SqlDataReader reader)
        {
            _discountName = reader.GetOrdinal("Discount Name");
            _discountType = reader.GetOrdinal("Discount Type");
            _discountAmount = reader.GetOrdinal("Discount Amount");
            _accessLevel = reader.GetOrdinal("Access Level");
            _categoryLink = reader.GetOrdinal("Category Link");
            _menuItemLink = reader.GetOrdinal("Menu Item Link");
            _maxDiscountAmount = reader.GetOrdinal("Max Discount Amount");
            _match = reader.GetOrdinal("Match");
            _taxable = reader.GetOrdinal("Taxable");
            _allowOnlyOne = reader.GetOrdinal("Allow Only One");
            _afterPrint = reader.GetOrdinal("After Print");
            _departmentLink = reader.GetOrdinal("Department Link");
            _autoMatch = reader.GetOrdinal("Auto Match");
            _discDepartment = reader.GetOrdinal("Disc Department");
            _type = reader.GetOrdinal("Type");
            _searchNumber = reader.GetOrdinal("Search Number");
            _startTime = reader.GetOrdinal("Start Time");
            _endTime = reader.GetOrdinal("End Time");
            _autoDeduct = reader.GetOrdinal("Auto Deduct");
            _tax1 = reader.GetOrdinal("Tax 1");
            _tax2 = reader.GetOrdinal("Tax 2");
            _expires = reader.GetOrdinal("Expires");
            _discountId = reader.GetOrdinal("Discount Id");
            _tracking = reader.GetOrdinal("Tracking");
        }

        public DiscountDto MapRow(SqlDataReader reader) => new()
        {
            DiscountName = GetString(reader, _discountName),
            DiscountType = GetString(reader, _discountType),
            DiscountAmount = GetDecimal(reader, _discountAmount),
            AccessLevel = GetInt(reader, _accessLevel),
            CategoryLink = GetString(reader, _categoryLink),
            MenuItemLink = GetString(reader, _menuItemLink),
            MaxDiscountAmount = GetDecimal(reader, _maxDiscountAmount),
            Match = GetString(reader, _match),
            Taxable = GetString(reader, _taxable),
            AllowOnlyOne = GetString(reader, _allowOnlyOne),
            AfterPrint = GetString(reader, _afterPrint),
            DepartmentLink = GetString(reader, _departmentLink),
            AutoMatch = GetString(reader, _autoMatch),
            DiscDepartment = GetString(reader, _discDepartment),
            Type = GetString(reader, _type),
            SearchNumber = GetString(reader, _searchNumber),
            StartTime = GetString(reader, _startTime),
            EndTime = GetString(reader, _endTime),
            AutoDeduct = GetString(reader, _autoDeduct),
            Tax1 = GetString(reader, _tax1),
            Tax2 = GetString(reader, _tax2),
            Expires = GetString(reader, _expires),
            DiscountId = GetString(reader, _discountId),
            Tracking = GetString(reader, _tracking)
        };

        private static string GetString(SqlDataReader reader, int ordinal) =>
            reader.IsDBNull(ordinal) ? string.Empty : reader.GetValue(ordinal)?.ToString() ?? string.Empty;

        private static int GetInt(SqlDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal)) return 0;
            return int.TryParse(reader.GetValue(ordinal)?.ToString(), out var value) ? value : 0;
        }

        private static decimal GetDecimal(SqlDataReader reader, int ordinal)
        {
            if (reader.IsDBNull(ordinal)) return 0m;
            return decimal.TryParse(reader.GetValue(ordinal)?.ToString(), out var value) ? value : 0m;
        }
    }
}

public sealed class DiscountDto
{
    public string DiscountName { get; init; } = "";
    public string DiscountType { get; init; } = "";
    public decimal DiscountAmount { get; init; }
    public int AccessLevel { get; init; }
    public string CategoryLink { get; init; } = "";
    public string MenuItemLink { get; init; } = "";
    public decimal MaxDiscountAmount { get; init; }
    public string Match { get; init; } = "";
    public string Taxable { get; init; } = "";
    public string AllowOnlyOne { get; init; } = "";
    public string AfterPrint { get; init; } = "";
    public string DepartmentLink { get; init; } = "";
    public string AutoMatch { get; init; } = "";
    public string DiscDepartment { get; init; } = "";
    public string Type { get; init; } = "";
    public string SearchNumber { get; init; } = "";
    public string StartTime { get; init; } = "";
    public string EndTime { get; init; } = "";
    public string AutoDeduct { get; init; } = "";
    public string Tax1 { get; init; } = "";
    public string Tax2 { get; init; } = "";
    public string Expires { get; init; } = "";
    public string DiscountId { get; init; } = "";
    public string Tracking { get; init; } = "";
}