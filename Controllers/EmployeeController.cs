using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    [ApiController]
    [Route("api/sync/employees")]
    public class EmployeeSyncController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public EmployeeSyncController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet("changes")]
        public IActionResult GetEmployeeChanges([FromQuery] int storeId, [FromQuery] DateTime? since)
        {
            try
            {
                string cs = _configuration.GetConnectionString("DefaultConnection");
                var changes = new List<EmployeeChangeDto>();

                using var cn = new SqlConnection(cs);
                cn.Open();

                string sql = @"
SELECT 
    Id,
    FirstName,
    LastName,
    Pin,
    RoleName,
    JobCode,
    IsActive,
    Email,
    Phone,
    ModifiedUtc
FROM Employees
WHERE (@since IS NULL OR ModifiedUtc > @since)
ORDER BY ModifiedUtc, Id";

                using var cmd = new SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@since", (object?)since ?? DBNull.Value);

                using var rdr = cmd.ExecuteReader();

                long changeId = 1;

                while (rdr.Read())
                {
                    string employeeId = rdr["Id"]?.ToString() ?? "";

                    string firstName = rdr["FirstName"]?.ToString() ?? "";
                    string lastName = rdr["LastName"]?.ToString() ?? "";
                    string displayName = (firstName + " " + lastName).Trim();

                    var payload = new EmployeePayloadDto
                    {
                        EmployeeId = employeeId,
                        FirstName = firstName,
                        LastName = lastName,
                        DisplayName = displayName,
                        PIN = rdr["Pin"]?.ToString() ?? "",
                        RoleName = rdr["RoleName"]?.ToString() ?? "",
                        JobCode = rdr["JobCode"]?.ToString() ?? "",
                        IsActive = rdr["IsActive"] != DBNull.Value && Convert.ToBoolean(rdr["IsActive"]),
                        Email = rdr["Email"]?.ToString() ?? "",
                        Phone = rdr["Phone"]?.ToString() ?? "",
                        ModifiedUtc = rdr["ModifiedUtc"] == DBNull.Value
                            ? null
                            : Convert.ToDateTime(rdr["ModifiedUtc"]).ToUniversalTime()
                    };

                    changes.Add(new EmployeeChangeDto
                    {
                        ChangeId = changeId++,
                        EntityType = "Employee",
                        EntityId = employeeId,
                        Action = payload.IsActive ? "Upsert" : "Deactivate",
                        ModifiedUtc = payload.ModifiedUtc,
                        Payload = payload
                    });
                }

                var result = new EmployeeSyncResponseDto
                {
                    Success = true,
                    ServerUtc = DateTime.UtcNow,
                    Changes = changes
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error: " + ex.Message);
            }
        }
    }

    public class EmployeeSyncResponseDto
    {
        public bool Success { get; set; }
        public DateTime? ServerUtc { get; set; }
        public List<EmployeeChangeDto> Changes { get; set; } = new();
    }

    public class EmployeeChangeDto
    {
        public long ChangeId { get; set; }
        public string EntityType { get; set; } = "";
        public string EntityId { get; set; } = "";
        public string Action { get; set; } = "";
        public DateTime? ModifiedUtc { get; set; }
        public EmployeePayloadDto? Payload { get; set; }
    }

    public class EmployeePayloadDto
    {
        public string EmployeeId { get; set; } = "";
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public string PIN { get; set; } = "";
        public string RoleName { get; set; } = "";
        public string JobCode { get; set; } = "";
        public bool IsActive { get; set; }
        public string Email { get; set; } = "";
        public string Phone { get; set; } = "";
        public DateTime? ModifiedUtc { get; set; }
    }
}