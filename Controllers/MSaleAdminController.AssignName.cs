using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    // Account Maintenance ("Assign Name") — house / charge account management.
    // Backs Views/MSaleAdmin/AssignName.cshtml against [HouseAccountData]
    // plus the [Starting Balance], [HouseCharges] and [HousePayments] tables.
    public partial class MSaleAdminController
    {
        private SqlConnection AnConn() =>
            new SqlConnection(_configuration.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public IActionResult AssignName() =>
            View("~/Views/MSaleAdmin/AssignName.cshtml");

        // ── Lookup lists ───────────────────────────────────────────────────
        [HttpGet]
        public IActionResult GetAccountNames()
        {
            using var cn = AnConn();
            var rows = cn.Query<AccountNameDto>(@"
                SELECT [Scan Name] AS ScanName, ISNULL([Acct Number],'') AS AcctNumber
                FROM [HouseAccountData]
                WHERE [Scan Name] IS NOT NULL AND [Scan Name] <> ''
                ORDER BY [Scan Name]");
            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetAccountTypes()
        {
            using var cn = AnConn();
            var rows = cn.Query<AccountTypeDto>(@"
                SELECT DISTINCT [Account Type] AS Name
                FROM [HouseAccountData]
                WHERE [Account Type] IS NOT NULL AND [Account Type] <> ''
                ORDER BY [Account Type]");
            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetDiscounts()
        {
            using var cn = AnConn();
            try
            {
                var names = cn.Query<string>(
                    "SELECT [Discount Name] FROM [Discounts] ORDER BY [Discount Name]");
                return Json(names);
            }
            catch
            {
                return Json(new List<string>());
            }
        }

        // ── Load one account (by name, else by number) ─────────────────────
        [HttpGet]
        public IActionResult GetAccount(string? scanName, string? acctNumber)
        {
            using var cn = AnConn();

            bool byName = !string.IsNullOrWhiteSpace(scanName);
            string where = byName ? "[Scan Name] = @K" : "[Acct Number] = @K";
            string key = byName ? scanName! : (acctNumber ?? "");

            var acct = cn.QueryFirstOrDefault<AccountDto>($@"
                SELECT TOP 1
                    [Scan Name]                            AS ScanName,
                    ISNULL([Acct Number],'')               AS AcctNumber,
                    ISNULL([Account Type],'House Account') AS AccountType,
                    ISNULL([Discount],'None')              AS Discount,
                    ISNULL([Address],'')                   AS Address,
                    ISNULL([City],'')                      AS City,
                    ISNULL([State],'')                     AS State,
                    ISNULL([Zip],'')                       AS Zip,
                    ISNULL([Phone],'')                     AS Phone,
                    ISNULL([Cell],'')                      AS Cell,
                    ISNULL([EMail],'')                     AS Email,
                    ISNULL([Fax],'')                       AS Fax,
                    ISNULL([Birthday],'')                  AS Birthday,
                    ISNULL([Anniversary],'')               AS Anniversary,
                    ISNULL([Account Expires On],'')        AS AccountExpiresOn,
                    ISNULL([Memo],'')                      AS Memo,
                    ISNULL([Max Daily Amount],0)           AS MaxDailyAmount,
                    ISNULL([Max Weekly Amount],0)          AS MaxWeeklyAmount,
                    ISNULL([Meal Limit],0)                 AS MealLimit,
                    ISNULL([Spc Interest 1],'')            AS SpcInterest1,
                    ISNULL([Spc Interest 2],'')            AS SpcInterest2,
                    ISNULL([Spc Interest 3],'')            AS SpcInterest3,
                    ISNULL([Spc Interest 4],'')            AS SpcInterest4,
                    ISNULL([Suspend],0)                    AS Suspend,
                    ISNULL([Delete Record],0)              AS DeleteRecord,
                    ISNULL([Credit Balance],0)             AS CreditBalance
                FROM [HouseAccountData]
                WHERE {where}", new { K = key });

            if (acct == null) return Json(null);

            // Balances are keyed by account (scan) name, per the original form.
            string name = acct.ScanName ?? "";

            decimal beginning = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Beginning Balance]) FROM [Starting Balance] WHERE [Account Name] = @N",
                new { N = name }) ?? 0m;

            decimal charges = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Adj Total]) FROM [HouseCharges] WHERE [Charged To] = @N AND [Voided] = 0",
                new { N = name }) ?? 0m;

            decimal payments = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Amount]) FROM [HousePayments] WHERE [Account Name] = @N",
                new { N = name }) ?? 0m;

            return Json(new
            {
                acct.ScanName, acct.AcctNumber, acct.AccountType, acct.Discount,
                acct.Address, acct.City, acct.State, acct.Zip, acct.Phone, acct.Cell,
                acct.Email, acct.Fax, acct.Birthday, acct.Anniversary, acct.AccountExpiresOn,
                acct.Memo, acct.MaxDailyAmount, acct.MaxWeeklyAmount, acct.MealLimit,
                acct.SpcInterest1, acct.SpcInterest2, acct.SpcInterest3, acct.SpcInterest4,
                acct.Suspend, acct.DeleteRecord, acct.CreditBalance,
                BeginningBalance = beginning,
                CurrentCharges = charges,
                Payments = payments,
                CurrentBalance = beginning + charges - payments
            });
        }

        // ── Create / update an account (upsert on Acct Number) ──────────────
        [HttpPost]
        public IActionResult SaveAccount([FromBody] AccountDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.ScanName))
                return Json(new { message = "Enter an account name." });
            if (string.IsNullOrWhiteSpace(dto.AcctNumber))
                return Json(new { message = "Must have an account number." });

            using var cn = AnConn();
            cn.Open();

            // Guard: same account number must not belong to a different name.
            var dupName = cn.ExecuteScalar<string>(
                "SELECT TOP 1 [Scan Name] FROM [HouseAccountData] WHERE [Acct Number] = @A",
                new { A = dto.AcctNumber.Trim() });
            if (dupName != null &&
                !string.Equals(dupName.Trim(), dto.ScanName.Trim(), StringComparison.OrdinalIgnoreCase))
                return Json(new { message = "Duplicate Account Number — cannot save." });

            var p = new
            {
                N = dto.ScanName.Trim(),
                A = dto.AcctNumber.Trim(),
                Type = string.IsNullOrWhiteSpace(dto.AccountType) ? "House Account" : dto.AccountType,
                Disc = string.IsNullOrWhiteSpace(dto.Discount) ? "None" : dto.Discount,
                Ad = dto.Address ?? "", Ci = dto.City ?? "", St = dto.State ?? "",
                Z = dto.Zip ?? "", Ph = dto.Phone ?? "", Ce = dto.Cell ?? "",
                Em = dto.Email ?? "", Fx = dto.Fax ?? "",
                Bd = dto.Birthday ?? "", An = dto.Anniversary ?? "", Ex = dto.AccountExpiresOn ?? "",
                Mo = dto.Memo ?? "",
                Md = dto.MaxDailyAmount, Mw = dto.MaxWeeklyAmount, Ml = dto.MealLimit,
                I1 = dto.SpcInterest1 ?? "", I2 = dto.SpcInterest2 ?? "",
                I3 = dto.SpcInterest3 ?? "", I4 = dto.SpcInterest4 ?? "",
                Su = dto.Suspend, Dl = dto.DeleteRecord, Cb = dto.CreditBalance
            };

            var exists = cn.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM [HouseAccountData] WHERE [Acct Number] = @A", new { A = p.A });

            if (exists > 0)
            {
                cn.Execute(@"UPDATE [HouseAccountData] SET
                    [Scan Name]=@N,[Account Type]=@Type,[Discount]=@Disc,
                    [Address]=@Ad,[City]=@Ci,[State]=@St,[Zip]=@Z,[Phone]=@Ph,[Cell]=@Ce,
                    [EMail]=@Em,[Fax]=@Fx,[Birthday]=@Bd,[Anniversary]=@An,[Account Expires On]=@Ex,
                    [Memo]=@Mo,[Max Daily Amount]=@Md,[Max Weekly Amount]=@Mw,[Meal Limit]=@Ml,
                    [Spc Interest 1]=@I1,[Spc Interest 2]=@I2,[Spc Interest 3]=@I3,[Spc Interest 4]=@I4,
                    [Suspend]=@Su,[Delete Record]=@Dl,[Credit Balance]=@Cb
                    WHERE [Acct Number]=@A", p);
            }
            else
            {
                cn.Execute(@"INSERT INTO [HouseAccountData]
                    ([Scan Name],[Acct Number],[Account Type],[Discount],
                     [Address],[City],[State],[Zip],[Phone],[Cell],[EMail],[Fax],
                     [Birthday],[Anniversary],[Account Expires On],[Memo],
                     [Max Daily Amount],[Max Weekly Amount],[Meal Limit],
                     [Spc Interest 1],[Spc Interest 2],[Spc Interest 3],[Spc Interest 4],
                     [Suspend],[Delete Record],[Credit Balance])
                    VALUES
                    (@N,@A,@Type,@Disc,@Ad,@Ci,@St,@Z,@Ph,@Ce,@Em,@Fx,
                     @Bd,@An,@Ex,@Mo,@Md,@Mw,@Ml,@I1,@I2,@I3,@I4,@Su,@Dl,@Cb)", p);
            }

            return Json(new { message = "Saved....." });
        }

        // ── Remove an account ───────────────────────────────────────────────
        [HttpPost]
        public IActionResult DeleteAccount(string? acctNumber, string? scanName)
        {
            using var cn = AnConn();
            int n;
            if (!string.IsNullOrWhiteSpace(acctNumber))
                n = cn.Execute("DELETE FROM [HouseAccountData] WHERE [Acct Number] = @A", new { A = acctNumber });
            else if (!string.IsNullOrWhiteSpace(scanName))
                n = cn.Execute("DELETE FROM [HouseAccountData] WHERE [Scan Name] = @N", new { N = scanName });
            else
                return Json(new { message = "Nothing to remove." });

            return Json(new { message = n > 0 ? "Account removed." : "Account not found." });
        }

        // ── Account Activity report ─────────────────────────────────────────
        [HttpGet]
        public IActionResult AccountActivity() =>
            View("~/Views/MSaleAdmin/AccountActivity.cshtml");

        [HttpGet]
        public IActionResult GetAccountActivity(string? scanName, string? startDate, string? endDate)
        {
            if (string.IsNullOrWhiteSpace(scanName))
                return Json(new { error = "Select an account." });

            DateTime start = DateTime.TryParse(startDate, out var s) ? s.Date : new DateTime(DateTime.Today.Year, 1, 1);
            DateTime end = DateTime.TryParse(endDate, out var e) ? e.Date : DateTime.Today;

            using var cn = AnConn();

            // Opening balance = seeded starting balance + activity before the start date.
            decimal seeded = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Beginning Balance]) FROM [Starting Balance] WHERE [Account Name] = @N",
                new { N = scanName }) ?? 0m;
            decimal priorCharges = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Adj Total]) FROM [HouseCharges] WHERE [Charged To] = @N AND [Voided] = 0 AND [Check Date] < @Start",
                new { N = scanName, Start = start }) ?? 0m;
            decimal priorPayments = cn.ExecuteScalar<decimal?>(
                "SELECT SUM([Amount]) FROM [HousePayments] WHERE [Account Name] = @N AND TRY_CAST([Date] AS date) < @Start",
                new { N = scanName, Start = start }) ?? 0m;
            decimal opening = seeded + priorCharges - priorPayments;

            var charges = cn.Query<ActivityRaw>(@"
                SELECT [Check Date] AS [Date], [Check Number] AS Reference,
                       [Adj Total] AS Amount, 'Charge' AS Kind, '' AS Info
                FROM [HouseCharges]
                WHERE [Charged To] = @N AND [Voided] = 0 AND [Check Date] BETWEEN @Start AND @End",
                new { N = scanName, Start = start, End = end });

            var payments = cn.Query<ActivityRaw>(@"
                SELECT TRY_CAST([Date] AS date) AS [Date], [Check Number] AS Reference,
                       [Amount] AS Amount, 'Payment' AS Kind, ISNULL([Payment Type],'') AS Info
                FROM [HousePayments]
                WHERE [Account Name] = @N AND TRY_CAST([Date] AS date) BETWEEN @Start AND @End",
                new { N = scanName, Start = start, End = end });

            // Merge in date order, charges before payments on the same day, and run the balance.
            var ordered = charges.Concat(payments)
                .OrderBy(r => r.Date ?? DateTime.MinValue)
                .ThenBy(r => r.Kind == "Charge" ? 0 : 1)
                .ThenBy(r => r.Reference)
                .ToList();

            decimal running = opening, totalCharges = 0m, totalPayments = 0m;
            var rows = new List<object>();
            foreach (var r in ordered)
            {
                bool isCharge = r.Kind == "Charge";
                if (isCharge) { running += r.Amount; totalCharges += r.Amount; }
                else { running -= r.Amount; totalPayments += r.Amount; }

                rows.Add(new
                {
                    date = r.Date?.ToString("MM/dd/yyyy") ?? "",
                    kind = r.Kind,
                    reference = r.Reference,
                    info = r.Info ?? "",
                    charge = isCharge ? r.Amount : 0m,
                    payment = isCharge ? 0m : r.Amount,
                    balance = running
                });
            }

            return Json(new
            {
                scanName,
                startDate = start.ToString("yyyy-MM-dd"),
                endDate = end.ToString("yyyy-MM-dd"),
                openingBalance = opening,
                totalCharges,
                totalPayments,
                endingBalance = opening + totalCharges - totalPayments,
                rows
            });
        }
    }

    // ── DTOs ────────────────────────────────────────────────────────────────
    public class ActivityRaw
    {
        public DateTime? Date { get; set; }
        public int Reference { get; set; }
        public decimal Amount { get; set; }
        public string Kind { get; set; } = "";
        public string Info { get; set; } = "";
    }

    public class AccountNameDto
    {
        public string ScanName { get; set; } = "";
        public string AcctNumber { get; set; } = "";
    }

    public class AccountTypeDto
    {
        public string Name { get; set; } = "";
    }

    public class AccountDto
    {
        public string ScanName { get; set; } = "";
        public string AcctNumber { get; set; } = "";
        public string AccountType { get; set; } = "House Account";
        public string Discount { get; set; } = "None";
        public string Address { get; set; } = "";
        public string City { get; set; } = "";
        public string State { get; set; } = "";
        public string Zip { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Cell { get; set; } = "";
        public string Email { get; set; } = "";
        public string Fax { get; set; } = "";
        public string Birthday { get; set; } = "";
        public string Anniversary { get; set; } = "";
        public string AccountExpiresOn { get; set; } = "";
        public string Memo { get; set; } = "";
        public decimal MaxDailyAmount { get; set; }
        public decimal MaxWeeklyAmount { get; set; }
        public int MealLimit { get; set; }
        public string SpcInterest1 { get; set; } = "";
        public string SpcInterest2 { get; set; } = "";
        public string SpcInterest3 { get; set; } = "";
        public string SpcInterest4 { get; set; } = "";
        public int Suspend { get; set; }
        public int DeleteRecord { get; set; }
        public int CreditBalance { get; set; }
    }
}
