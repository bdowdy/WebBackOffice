using Microsoft.AspNetCore.Mvc;
using MSaleWebServer.Models;
using System.Configuration;

namespace MSaleWebServer.Controllers
{
    public class TimeMessages(IConfiguration configuration) : Controller
    {
        private readonly IConfiguration _configuration = configuration;

        public IActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetTimeMessageEmployees()
        {
            try
            {
                var list = new List<string>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT 
                RTRIM(ISNULL([First Name], '')) + ' ' + RTRIM(ISNULL([Last Name], '')) AS EmployeeName
            FROM [dbo].[General Data]
            WHERE ISNULL([Terminated], 0) = 0
            ORDER BY [Last Name], [First Name]";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd["EmployeeName"].ToString() ?? "");

                return Json(new { success = true, data = list });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetTimeMessageDepartments()
        {
            try
            {
                var list = new List<string>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT [Job Title]
            FROM [dbo].[Job Titles]
            ORDER BY [Job Title]";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd[0].ToString() ?? "");

                return Json(new { success = true, data = list });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetUndeliveredMessageNames()
        {
            try
            {
                var list = new List<string>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT DISTINCT [Name]
            FROM [dbo].[Message All]
            WHERE ISNULL([Name], '') <> ''
            ORDER BY [Name]";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                    list.Add(rd[0].ToString() ?? "");

                return Json(new { success = true, data = list });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public IActionResult GetMessagesForEmployee(string name)
        {
            try
            {
                var list = new List<object>();
                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                string sql = @"
            SELECT [Date], [Message], [Message from]
            FROM [dbo].[Message All]
            WHERE [Name] = @Name
            ORDER BY TRY_CAST([Date] AS datetime) DESC";

                using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                cmd.Parameters.AddWithValue("@Name", name ?? "");

                using var rd = cmd.ExecuteReader();

                while (rd.Read())
                {
                    list.Add(new
                    {
                        date = rd["Date"].ToString(),
                        message = rd["Message"].ToString(),
                        messageFrom = rd["Message from"].ToString()
                    });
                }

                return Json(new { success = true, data = list });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult SaveTimeMessages([FromBody] TimeMessageDto item)
        {
            try
            {
                if (item.Names == null || item.Names.Count == 0)
                    return Json(new { success = false, message = "Select at least one employee." });

                if (string.IsNullOrWhiteSpace(item.Message))
                    return Json(new { success = false, message = "Enter a message." });

                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                foreach (var name in item.Names)
                {
                    string sql = @"
                INSERT INTO [dbo].[Message All]
                (
                    [Date],
                    [Name],
                    [Message],
                    [Message from]
                )
                VALUES
                (
                    @Date,
                    @Name,
                    @Message,
                    @MessageFrom
                )";

                    using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                    cmd.Parameters.AddWithValue("@Date", item.DeliveryDate ?? DateTime.Today.ToString("MM/dd/yyyy"));
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.Parameters.AddWithValue("@Message", item.Message);
                    cmd.Parameters.AddWithValue("@MessageFrom", item.MessageFrom ?? "MSaleWebServer");

                    cmd.ExecuteNonQuery();
                }

                return Json(new { success = true, message = "Message recorded." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public IActionResult DeleteTimeMessages([FromBody] TimeMessageDto item)
        {
            try
            {
                if (item.Names == null || item.Names.Count == 0)
                    return Json(new { success = false, message = "Select an employee." });

                var cs = _configuration.GetConnectionString("DefaultConnection");

                using var cn = new Microsoft.Data.SqlClient.SqlConnection(cs);
                cn.Open();

                foreach (var name in item.Names)
                {
                    string sql = "DELETE FROM [dbo].[Message All] WHERE [Name] = @Name";

                    using var cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, cn);
                    cmd.Parameters.AddWithValue("@Name", name);
                    cmd.ExecuteNonQuery();
                }

                return Json(new { success = true, message = "Message deleted." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }


    }
}
