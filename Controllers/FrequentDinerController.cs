using Dapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MSaleWebServer.Controllers
{
    public class FrequentDinerController : Controller
    {
        private readonly IConfiguration _cfg;
        public FrequentDinerController(IConfiguration cfg) => _cfg = cfg;

        SqlConnection Cn() => new(_cfg.GetConnectionString("DefaultConnection"));

        [HttpGet]
        public IActionResult Index() =>
            View("~/Views/MSaleAdmin/FrequentDiner.cshtml");

        // â”€â”€ CUSTOMERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        [HttpGet]
        public IActionResult GetCustomers(string? search)
        {
            using var cn = Cn();
            string sql = @"
                SELECT [Scan Name] AS ScanName, [Acct Number] AS AcctNumber,
                       ISNULL([Address],'') AS Address, ISNULL([City],'') AS City,
                       ISNULL([State],'') AS State, ISNULL([zip],'') AS Zip,
                       ISNULL([Phone],'') AS Phone, ISNULL([EMail],'') AS Email,
                       ISNULL([Birthday],'') AS Birthday, ISNULL([Cell],'') AS Cell,
                       ISNULL([Suspend],0) AS Suspend,
                       ISNULL([Credit Balance],0) AS CreditBalance,
                       ISNULL([Spc Interest 1],'') AS Interest1,
                       ISNULL([Spc Interest 2],'') AS Interest2
                FROM [MSaleData].[dbo].[HouseAccountData]
                WHERE [Account Type] = 'Frequent Diner'";
            if (!string.IsNullOrWhiteSpace(search))
                sql += @" AND ([Scan Name] LIKE @S OR [Acct Number] LIKE @S
                              OR [Phone] LIKE @S OR [EMail] LIKE @S)";
            sql += " ORDER BY [Scan Name]";
            var rows = cn.Query(sql, new { S = "%" + search + "%" });
            return Json(rows);
        }

        [HttpGet]
        public IActionResult GetCustomer(string acctNumber)
        {
            using var cn = Cn();
            var cust = cn.QueryFirstOrDefault(@"
                SELECT [Scan Name] AS ScanName, [Acct Number] AS AcctNumber,
                       ISNULL([Address],'') AS Address, ISNULL([City],'') AS City,
                       ISNULL([State],'') AS State, ISNULL([zip],'') AS Zip,
                       ISNULL([Phone],'') AS Phone, ISNULL([EMail],'') AS Email,
                       ISNULL([Birthday],'') AS Birthday, ISNULL([Cell],'') AS Cell,
                       ISNULL([Anniversary],'') AS Anniversary,
                       ISNULL([Suspend],0) AS Suspend,
                       ISNULL([Credit Balance],0) AS CreditBalance,
                       ISNULL([Spc Interest 1],'') AS Interest1,
                       ISNULL([Spc Interest 2],'') AS Interest2,
                       ISNULL([Spc Interest 3],'') AS Interest3,
                       ISNULL([Spc Interest 4],'') AS Interest4
                FROM [MSaleData].[dbo].[HouseAccountData]
                WHERE [Acct Number] = @A", new { A = acctNumber });
            return Json(cust);
        }

        [HttpPost]
        public IActionResult SaveCustomer([FromBody] CustomerDto dto)
        {
            using var cn = Cn(); cn.Open();
            // duplicate account check
            var dup = cn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM [MSaleData].[dbo].[HouseAccountData] WHERE [Acct Number]=@A AND [Scan Name]<>@N",
                new { A = dto.AcctNumber, N = dto.ScanName });
            if (dup > 0) return BadRequest(new { error = "Account number already used by a different customer." });

            var exists = cn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM [MSaleData].[dbo].[HouseAccountData] WHERE [Acct Number]=@A",
                new { A = dto.AcctNumber });

            if (exists > 0)
            {
                cn.Execute(@"UPDATE [MSaleData].[dbo].[HouseAccountData] SET
                    [Scan Name]=@N,[Address]=@Ad,[City]=@Ci,[State]=@St,[zip]=@Z,
                    [Phone]=@Ph,[EMail]=@Em,[Birthday]=@Bd,[Cell]=@Ce,
                    [Anniversary]=@An,[Suspend]=@Su,[Spc Interest 1]=@I1,
                    [Spc Interest 2]=@I2,[Spc Interest 3]=@I3,[Spc Interest 4]=@I4
                    WHERE [Acct Number]=@A",
                    new { N=dto.ScanName,Ad=dto.Address,Ci=dto.City,St=dto.State,Z=dto.Zip,
                          Ph=dto.Phone,Em=dto.Email,Bd=dto.Birthday,Ce=dto.Cell,
                          An=dto.Anniversary,Su=dto.Suspend,I1=dto.Interest1,
                          I2=dto.Interest2,I3=dto.Interest3,I4=dto.Interest4,A=dto.AcctNumber });
            }
            else
            {
                cn.Execute(@"INSERT INTO [MSaleData].[dbo].[HouseAccountData]
                    ([Scan Name],[Acct Number],[Account Type],[Address],[City],[State],[zip],
                     [Phone],[EMail],[Birthday],[Cell],[Anniversary],[Suspend],
                     [Spc Interest 1],[Spc Interest 2],[Spc Interest 3],[Spc Interest 4])
                    VALUES(@N,@A,'Frequent Diner',@Ad,@Ci,@St,@Z,@Ph,@Em,@Bd,@Ce,@An,@Su,@I1,@I2,@I3,@I4)",
                    new { N=dto.ScanName,A=dto.AcctNumber,Ad=dto.Address,Ci=dto.City,St=dto.State,
                          Z=dto.Zip,Ph=dto.Phone,Em=dto.Email,Bd=dto.Birthday,Ce=dto.Cell,
                          An=dto.Anniversary,Su=dto.Suspend,I1=dto.Interest1,
                          I2=dto.Interest2,I3=dto.Interest3,I4=dto.Interest4 });
                // seed lifetime totals row
                cn.Execute(@"IF NOT EXISTS (SELECT 1 FROM [MSaleData].[dbo].[Non Reset Data] WHERE [Customer Account]=@A)
                    INSERT INTO [MSaleData].[dbo].[Non Reset Data]([Customer Account],[Customer Name],[Total Visits],[Total Dollars],[Total Points])
                    VALUES(@A,@N,0,0,0)", new { A=dto.AcctNumber, N=dto.ScanName });
            }
            return Ok();
        }

        [HttpPost]
        public IActionResult DeleteCustomer([FromBody] DeleteCustomerDto dto)
        {
            using var cn = Cn();
            cn.Execute("DELETE FROM [MSaleData].[dbo].[HouseAccountData] WHERE [Acct Number]=@A", new { A=dto.AcctNumber });
            cn.Execute("DELETE FROM [MSaleData].[dbo].[FrequentDinerDetail] WHERE [Customer Account]=@A", new { A=dto.AcctNumber });
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Non Reset Data] WHERE [Customer Account]=@A", new { A=dto.AcctNumber });
            return Ok();
        }

        [HttpPost]
        public IActionResult ChangeCustomerName([FromBody] NameChangeDto dto)
        {
            using var cn = Cn();
            cn.Execute("UPDATE [MSaleData].[dbo].[HouseAccountData] SET [Scan Name]=@New WHERE [Scan Name]=@Old AND [Acct Number]=@A",
                new { New=dto.NewName, Old=dto.OldName, A=dto.AcctNumber });
            cn.Execute("UPDATE [MSaleData].[dbo].[FrequentDinerDetail] SET [Customer Name]=@New WHERE [Customer Name]=@Old AND [Customer Account]=@A",
                new { New=dto.NewName, Old=dto.OldName, A=dto.AcctNumber });
            cn.Execute("UPDATE [MSaleData].[dbo].[Non Reset Data] SET [Customer Name]=@New WHERE [Customer Name]=@Old AND [Customer Account]=@A",
                new { New=dto.NewName, Old=dto.OldName, A=dto.AcctNumber });
            return Ok();
        }

        [HttpPost]
        public IActionResult ChangeAccountNumber([FromBody] AcctChangeDto dto)
        {
            using var cn = Cn();
            // check new number not already taken
            var dup = cn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM [MSaleData].[dbo].[HouseAccountData] WHERE [Acct Number]=@New",
                new { New=dto.NewAcct });
            if (dup > 0) return BadRequest(new { error = "New account number is already in use." });
            cn.Execute("UPDATE [MSaleData].[dbo].[HouseAccountData] SET [Acct Number]=@New WHERE [Acct Number]=@Old",
                new { New=dto.NewAcct, Old=dto.OldAcct });
            cn.Execute("UPDATE [MSaleData].[dbo].[FrequentDinerDetail] SET [Customer Account]=@New WHERE [Customer Account]=@Old",
                new { New=dto.NewAcct, Old=dto.OldAcct });
            cn.Execute("UPDATE [MSaleData].[dbo].[Non Reset Data] SET [Customer Account]=@New WHERE [Customer Account]=@Old",
                new { New=dto.NewAcct, Old=dto.OldAcct });
            return Ok();
        }

        // â”€â”€ CUSTOMER ACTIVITY â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        [HttpGet]
        public IActionResult GetCustomerActivity(string acctNumber)
        {
            using var cn = Cn();
            // per-plan totals
            var plans = cn.Query(@"
                SELECT [Plan Name] AS PlanName,
                       SUM([Price])  AS TotalDollars,
                       SUM([Points]) AS TotalPoints,
                       COUNT(DISTINCT [Check Number]) AS CheckCount
                FROM [MSaleData].[dbo].[FrequentDinerDetail]
                WHERE [Customer Account]=@A AND [Voided]=0
                GROUP BY [Plan Name]
                ORDER BY [Plan Name]", new { A=acctNumber });

            // lifetime totals
            var lifetime = cn.QueryFirstOrDefault(@"
                SELECT ISNULL([Total Visits],0) AS TotalVisits,
                       ISNULL([Total Dollars],0) AS TotalDollars,
                       ISNULL([Total Points],0)  AS TotalPoints
                FROM [MSaleData].[dbo].[Non Reset Data]
                WHERE [Customer Account]=@A", new { A=acctNumber });

            // recent transactions (last 50)
            var recent = cn.Query(@"
                SELECT DISTINCT TRY_CAST([Date] AS date) AS TxDate,
                       [Plan Name] AS PlanName, [Check Number] AS CheckNumber,
                       SUM([Price])  AS TPrice,
                       SUM([Points]) AS TPoints,
                       [Customer Name] AS CustomerName
                FROM [MSaleData].[dbo].[FrequentDinerDetail]
                WHERE [Customer Account]=@A AND [Voided]=0
                GROUP BY [Customer Name],[Customer Account],[Plan Name],[Check Number],[Date]
                ORDER BY TRY_CAST([Date] AS date) DESC", new { A=acctNumber });

            return Json(new { plans, lifetime, recent });
        }

        [HttpPost]
        public IActionResult AddManagerTransaction([FromBody] MgrTxDto dto)
        {
            using var cn = Cn();
            cn.Execute(@"INSERT INTO [MSaleData].[dbo].[FrequentDinerDetail]
                ([Date],[Time],[Item],[Voided],[Price],[Points],
                 [Customer Name],[Customer Account],[Check Number],[Quantity],[Selection Number],[Plan Name])
                VALUES(@D,@T,'Manager Adjustment',0,@Price,@Points,@N,@A,-100,1,-1,@Plan)",
                new { D=DateTime.Today.ToString("MM/dd/yyyy"),
                      T=DateTime.Now.ToString("hh:mm tt"),
                      Price=dto.Dollars, Points=dto.Points,
                      N=dto.CustomerName, A=dto.AcctNumber, Plan=dto.PlanName });

            // update lifetime totals
            cn.Execute(@"UPDATE [MSaleData].[dbo].[Non Reset Data]
                SET [Total Dollars] = ISNULL([Total Dollars],0) + @Price,
                    [Total Points]  = ISNULL([Total Points],0)  + @Points
                WHERE [Customer Account]=@A",
                new { Price=dto.Dollars, Points=dto.Points, A=dto.AcctNumber });
            return Ok();
        }

        // â”€â”€ PLANS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        [HttpGet]
        public IActionResult GetPlans()
        {
            using var cn = Cn();
            var plans = cn.Query(@"
                SELECT [Plan Name] AS PlanName,
                       ISNULL([Discount Name],'') AS DiscountName,
                       ISNULL([Qualify Amount],0) AS QualifyAmount,
                       ISNULL([Points],0) AS IsPoints,
                       ISNULL([Dollars],0) AS IsDollars,
                       ISNULL([Search By],'') AS SearchBy
                FROM [MSaleData].[dbo].[Plan Information]
                ORDER BY [Plan Name]");
            return Json(plans);
        }

        [HttpGet]
        public IActionResult GetPlan(string planName)
        {
            using var cn = Cn();
            var info = cn.QueryFirstOrDefault(@"
                SELECT [Plan Name] AS PlanName,
                       ISNULL([Discount Name],'') AS DiscountName,
                       ISNULL([Qualify Amount],0) AS QualifyAmount,
                       ISNULL([Points],0) AS IsPoints,
                       ISNULL([Dollars],0) AS IsDollars,
                       ISNULL([Search By],'') AS SearchBy
                FROM [MSaleData].[dbo].[Plan Information]
                WHERE [Plan Name]=@P", new { P=planName });

            var items = cn.Query(@"
                SELECT [Item Qualifier] AS ItemQualifier, ISNULL([Value],0) AS Value
                FROM [MSaleData].[dbo].[Plan Data]
                WHERE [Plan Name]=@P ORDER BY [Item Qualifier]", new { P=planName });

            var events = cn.Query(@"
                SELECT [Event Name] AS EventName,
                       ISNULL([Function Name],'Regular') AS FunctionName,
                       ISNULL([Dollar Amount],0) AS DollarAmount
                FROM [MSaleData].[dbo].[Plan Events]
                WHERE [Plan Name]=@P ORDER BY [Event Name]", new { P=planName });

            return Json(new { info, items, events });
        }

        [HttpPost]
        public IActionResult SavePlan([FromBody] PlanDto dto)
        {
            using var cn = Cn(); cn.Open();
            var exists = cn.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM [MSaleData].[dbo].[Plan Information] WHERE [Plan Name]=@P",
                new { P=dto.PlanName });

            if (exists > 0)
                cn.Execute(@"UPDATE [MSaleData].[dbo].[Plan Information] SET
                    [Discount Name]=@D,[Qualify Amount]=@Q,[Points]=@Pts,[Dollars]=@Dol,[Search By]=@S
                    WHERE [Plan Name]=@P",
                    new { D=dto.DiscountName,Q=dto.QualifyAmount,Pts=dto.IsPoints?1:0,
                          Dol=dto.IsDollars?1:0,S=dto.SearchBy,P=dto.PlanName });
            else
                cn.Execute(@"INSERT INTO [MSaleData].[dbo].[Plan Information]
                    ([Plan Name],[Discount Name],[Qualify Amount],[Points],[Dollars],[Search By])
                    VALUES(@P,@D,@Q,@Pts,@Dol,@S)",
                    new { P=dto.PlanName,D=dto.DiscountName,Q=dto.QualifyAmount,
                          Pts=dto.IsPoints?1:0,Dol=dto.IsDollars?1:0,S=dto.SearchBy });

            // replace plan data items
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Plan Data] WHERE [Plan Name]=@P", new { P=dto.PlanName });
            foreach (var item in dto.Items ?? [])
                cn.Execute(@"INSERT INTO [MSaleData].[dbo].[Plan Data]([Plan Name],[Item Qualifier],[Value],[Search By])
                    VALUES(@P,@I,@V,@S)",
                    new { P=dto.PlanName,I=item.ItemQualifier,V=item.Value,S=dto.SearchBy });

            // replace plan events
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Plan Events] WHERE [Plan Name]=@P", new { P=dto.PlanName });
            foreach (var ev in dto.Events ?? [])
                cn.Execute(@"INSERT INTO [MSaleData].[dbo].[Plan Events]([Plan Name],[Event Name],[Function Name],[Dollar Amount])
                    VALUES(@P,@E,@F,@D)",
                    new { P=dto.PlanName,E=ev.EventName,F=ev.FunctionName,D=ev.DollarAmount });

            return Ok();
        }

        [HttpPost]
        public IActionResult DeletePlan([FromBody] DeletePlanDto dto)
        {
            using var cn = Cn();
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Plan Information] WHERE [Plan Name]=@P", new { P=dto.PlanName });
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Plan Data] WHERE [Plan Name]=@P", new { P=dto.PlanName });
            cn.Execute("DELETE FROM [MSaleData].[dbo].[Plan Events] WHERE [Plan Name]=@P", new { P=dto.PlanName });
            return Ok();
        }

        // â”€â”€ REPORTS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        [HttpGet]
        public IActionResult GetCustomerStatus()
        {
            using var cn = Cn();
            // Per-customer, per-plan totals
            var activity = cn.Query(@"
                SELECT [Customer Name] AS CustomerName,
                       [Customer Account] AS AcctNumber,
                       [Plan Name] AS PlanName,
                       SUM([Price])  AS TotalDollars,
                       SUM([Points]) AS TotalPoints,
                       COUNT(DISTINCT [Check Number]) AS Visits
                FROM [MSaleData].[dbo].[FrequentDinerDetail]
                WHERE [Voided]=0
                GROUP BY [Customer Name],[Customer Account],[Plan Name]
                ORDER BY [Customer Name],[Plan Name]");

            // plan thresholds
            var thresholds = cn.Query(@"
                SELECT [Plan Name] AS PlanName,
                       ISNULL([Qualify Amount],0) AS QualifyAmount,
                       ISNULL([Points],0) AS IsPoints,
                       ISNULL([Dollars],0) AS IsDollars,
                       ISNULL([Discount Name],'') AS DiscountName
                FROM [MSaleData].[dbo].[Plan Information]");

            // lifetime totals
            var lifetime = cn.Query(@"
                SELECT [Customer Account] AS AcctNumber,
                       [Customer Name] AS CustomerName,
                       ISNULL([Total Visits],0)  AS TotalVisits,
                       ISNULL([Total Dollars],0) AS TotalDollars,
                       ISNULL([Total Points],0)  AS TotalPoints
                FROM [MSaleData].[dbo].[Non Reset Data]
                ORDER BY [Customer Name]");

            return Json(new { activity, thresholds, lifetime });
        }

        [HttpGet]
        public IActionResult GetCustomerList()
        {
            using var cn = Cn();
            var rows = cn.Query(@"
                SELECT [Scan Name] AS ScanName,
                       ISNULL([Acct Number],'') AS AcctNumber,
                       ISNULL([Phone],'') AS Phone,
                       ISNULL([Cell],'') AS Cell,
                       ISNULL([EMail],'') AS Email,
                       ISNULL([Address],'') AS Address,
                       ISNULL([City],'') AS City,
                       ISNULL([State],'') AS State,
                       ISNULL([zip],'') AS Zip,
                       ISNULL([Birthday],'') AS Birthday
                FROM [MSaleData].[dbo].[HouseAccountData]
                WHERE [Account Type] = 'Frequent Diner'
                ORDER BY [Scan Name]");
            return Json(rows);
        }

        // â”€â”€ LOOKUP DATA â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

        [HttpGet]
        public IActionResult GetDiscounts()
        {
            using var cn = Cn();
            try {
                var items = cn.Query<string>("SELECT [Discount Name] FROM [MSaleData].[dbo].[Discounts] ORDER BY [Discount Name]");
                return Json(items);
            } catch { return Json(new List<string>()); }
        }

        [HttpGet]
        public IActionResult GetMenuChoices(string type)
        {
            using var cn = Cn();
            try {
                string sql = type switch {
                    "Category"   => "SELECT [Category Name] AS Name FROM [MSaleData].[dbo].[Category Information] ORDER BY [Category Name]",
                    "Department" => "SELECT [Report Category] AS Name FROM [MSaleData].[dbo].[Report Categories] ORDER BY [Report Category]",
                    _            => "SELECT [Menu Item] AS Name FROM [MSaleData].[dbo].[Menu Items] ORDER BY [Menu Item]"
                };
                var names = cn.Query<string>(sql);
                return Json(names);
            } catch { return Json(new List<string>()); }
        }

        [HttpGet]
        public IActionResult GetTimedEvents()
        {
            using var cn = Cn();
            try {
                var events = cn.Query<string>(
                    "SELECT DISTINCT [Event Name] FROM [MSaleData].[dbo].[Timed Pricing] ORDER BY [Event Name]");
                return Json(events);
            } catch { return Json(new List<string>()); }
        }
    }

    // â”€â”€ DTOs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    public class CustomerDto {
        public string ScanName { get; set; } = "";
        public string AcctNumber { get; set; } = "";
        public string Address { get; set; } = "";
        public string City { get; set; } = "";
        public string State { get; set; } = "";
        public string Zip { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Email { get; set; } = "";
        public string Birthday { get; set; } = "";
        public string Cell { get; set; } = "";
        public string Anniversary { get; set; } = "";
        public int    Suspend { get; set; }
        public string Interest1 { get; set; } = "";
        public string Interest2 { get; set; } = "";
        public string Interest3 { get; set; } = "";
        public string Interest4 { get; set; } = "";
    }
    public class DeleteCustomerDto { public string AcctNumber { get; set; } = ""; }
    public class NameChangeDto { public string OldName { get; set; }=""; public string NewName { get; set; }=""; public string AcctNumber { get; set; }=""; }
    public class AcctChangeDto  { public string OldAcct { get; set; }=""; public string NewAcct { get; set; }=""; }
    public class MgrTxDto { public string AcctNumber{get;set;}=""; public string CustomerName{get;set;}=""; public string PlanName{get;set;}=""; public decimal Dollars{get;set;} public decimal Points{get;set;} }
    public class PlanDto {
        public string PlanName    { get; set; } = "";
        public string DiscountName{ get; set; } = "";
        public decimal QualifyAmount { get; set; }
        public bool IsPoints  { get; set; }
        public bool IsDollars { get; set; }
        public string SearchBy { get; set; } = "Menu Item";
        public List<PlanItemDto>?  Items  { get; set; }
        public List<PlanEventDto>? Events { get; set; }
    }
    public class PlanItemDto  { public string ItemQualifier{get;set;}=""; public decimal Value{get;set;} }
    public class PlanEventDto { public string EventName{get;set;}=""; public string FunctionName{get;set;}="Regular"; public decimal DollarAmount{get;set;} }
    public class DeletePlanDto { public string PlanName { get; set; } = ""; }
}
