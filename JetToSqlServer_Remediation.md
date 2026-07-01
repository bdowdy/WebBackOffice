# Jet → SQL Server SQL Remediation Guide — M$ (live POS app)

The `M$` app's databases now resolve to **SQL Server** (via `DatabaseNames.bas` →
`ReturnAdoConnection`), but many SQL strings still use **Access/Jet-only syntax**
that SQL Server rejects at runtime. This guide lists every site found and gives a
T-SQL-correct rewrite for each.

**Totals found (live `M$` folder only):**

| Construct | What | Sites |
|---|---|---|
| A | `#...#` date literals & `cdate()` in SQL | 328 |
| B | `Val([Col] & '')` used as a SQL function | 3 |
| C | `Format([Col],'HH:MM ampm')` in SQL | 2 |
| D | plain `cast([Col] as date/datetime/int)` in SQL | 309 |
| E | `&` concatenation between SQL columns/literals | 3 (same lines as B) |
| F | `Nz(` / `IIf(` inside SQL | 0 |

---

## Translation rules

| Jet / Access (breaks on SQL Server) | SQL Server (T-SQL) |
|---|---|
| `#" & someDate & "#"` (date literal) | `'" & Format(CDate(someDate),"yyyy-mm-dd") & "'"` (ISO, locale-safe) |
| `cdate([Col]) = #x#` | `TRY_CAST([Col] AS date) = 'x'` |
| `cdate([Col])` (order/compare) | `TRY_CAST([Col] AS date)` (or `AS datetime` / `AS time`) |
| `cast([Col] as date)` | `TRY_CAST([Col] AS date)` — `TRY_` so bad text → NULL, not an error |
| `cast([Col] as int)` (text column) | `TRY_CONVERT(int, [Col])` |
| `Val([Col] & '') = 0` | `ISNULL(TRY_CONVERT(int,[Col]),0) = 0` |
| `Format([Col],'HH:MM ampm')` | `TRY_CAST([Col] AS time)` (compare to a `TRY_CAST(... AS time)` value) |
| `[Col] & ''` (in SQL) | `ISNULL([Col],'')` |
| `IIf(a,b,c)` | `CASE WHEN a THEN b ELSE c END` |
| `Nz(x,y)` | `ISNULL(x,y)` |
| `datepart('ww', cdate([d]),2,2)` | `DATEPART(ISO_WEEK, TRY_CAST([d] AS date))` |
| `datepart('h',[Time])` | `DATEPART(hour, TRY_CAST([Time] AS time))` |

> **Mask trap:** VB `"HH:MM ampm"` means *minutes* (`MM`) in Jet but *month* in T-SQL.
> Never reuse that mask server-side. Build VB date strings as `"yyyy-mm-dd"` and time
> strings as `"HH:nn:ss"`.

### Suggested helpers (add to `functions.bas`)

Putting the ISO formatting in one place keeps the ~640 rewrites consistent and reduces
typos:

```vb
' Wrap a date as an ISO, quote-delimited SQL Server literal.
Public Function SqlDate(ByVal v As Variant) As String
    SqlDate = "'" & Format(CDate(v), "yyyy-mm-dd") & "'"
End Function

Public Function SqlDateTime(ByVal v As Variant) As String
    SqlDateTime = "'" & Format(CDate(v), "yyyy-mm-dd hh:nn:ss") & "'"
End Function

Public Function SqlTime(ByVal v As Variant) As String
    SqlTime = "'" & Format(CDate(v), "HH:nn:ss") & "'"
End Function
```

Then a rewrite reads cleanly, e.g.:

```vb
' BEFORE (Jet)
Sql = "Select * from [Sales] where cdate([Date]) = #" & CDate(DisplayDate) & "#"
' AFTER (SQL Server)
Sql = "Select * from [Sales] where TRY_CAST([Date] AS date) = " & SqlDate(DisplayDate)
```

---

## Construct A — `#...#` date delimiters and `cdate()` in SQL (328)

267 use `#...#` delimiters; 61 use `cdate([Col])` already quote-delimited.

**Representative rewrites:**

| file:line | original | rewrite |
|---|---|---|
| AuditDeposit.frm:560 | `where cdate([Date]) = #" & CDate(.TextMatrix(R,0)) & "#"` | `where TRY_CAST([Date] AS date) = " & SqlDate(.TextMatrix(R,0))` |
| AuditDeposit.frm:561 | `and cdate([Start Time]) = #" & CDate(.TextMatrix(R,1)) & "#"` | `and TRY_CAST([Start Time] AS time) = " & SqlTime(.TextMatrix(R,1))` |
| Closing_Report.frm:3193 | `where cdate([date]) between #" & a & "# and #" & b & "#"` | `where TRY_CAST([date] AS datetime) between " & SqlDateTime(a) & " and " & SqlDateTime(b)` |
| Closing_Report.frm:3944 | `delete from [Labor] where Cdate([date]) = #" & ClosingDate & "#"` | `delete from [Labor] where TRY_CAST([date] AS date) = " & SqlDate(ClosingDate)` |
| Display.frm:2267 | `from [Sales] where cdate([Date]) = #" & CDate(DisplayDate) & "#"` | `from [Sales] where TRY_CAST([Date] AS date) = " & SqlDate(DisplayDate)` |
| Display.frm:2436 | `where cdate([Date]) = #" & Format(CDate(SearchDate),"MM/DD/YYYY") & "#"` | `where TRY_CAST([Date] AS date) = " & SqlDate(SearchDate)` |
| Subs.bas:814 | `Where cdate([date]) Between #" & CDate(SDate) & " " & GetStartTime & "# and #" & SetEndingDate(EDate) & " " & SetEndingTime & "#"` | `Where TRY_CAST([date] AS datetime) Between '" & Format(CDate(SDate),"yyyy-mm-dd") & " " & GetStartTime & "' and '" & Format(CDate(SetEndingDate(EDate)),"yyyy-mm-dd") & " " & SetEndingTime & "'"` |

**`datepart` sub-case** (Jet quoted interval + inner `cdate`, mostly Display.frm):
4689, 4690, 4721, 4722, 6764, 6981, 6989 → `DATEPART(ISO_WEEK, TRY_CAST([date] AS date))` (or `week` per your rule).
Subs.bas:7386, 7391, 7392 → `DATEPART(hour, TRY_CAST([Time] AS time))`.

**All remaining `#...#` sites (file:line):**
Display.frm: 2652, 2684, 2699, 2713, 2725, 2752, 2767, 2787, 2883, 2884, 2901, 2902, 2916, 2927, 2928, 2981, 3037, 3266, 3279, 3280, 3283, 3284, 3393, 3394, 3417, 3418, 3439, 3440, 3462, 3463, 3481, 3658, 3659, 3685, 3686, 3698, 3699, 3701, 3728, 3729, 3750, 3751, 3774, 3775, 3845, 3846, 3915, 3916, 3942, 3943, 3952, 3953, 3976, 3977, 4079, 4080, 4097, 4098, 4123, 4124, 4165, 4269, 4281, 4382, 4383, 4415, 4416, 4437, 4438, 4673, 4719, 4760, 4793, 4822, 4877, 4912, 4948, 4959, 4976, 5001, 5026, 5253, 5254, 5308, 5309, 5373, 5374, 5378, 5379, 5477, 5478, 5885, 5886, 6029, 6128, 6129, 6133, 6134, 6153, 6154, 6158, 6159, 6200, 6201, 6387, 6388, 6394, 6395, 6401, 6402, 6524, 6525, 6557, 6587, 6613, 6645, 6780, 7226, 7229, 7232, 7235, 7238, 7241, 7244, 7247, 7250, 7253, 7256, 7259, 7262, 7267, 7270, 7766;
Financal_Reports.frm: 9119, 9718, 9722, 9730, 9928, 9945, 10345, 10386, 10633, 10634, 10882, 11068, 11139, 11178, 11209, 11239, 11880, 12019, 12165, 12175, 12314, 12315;
ListCards.frm: 2821, 3105; Mix.frm: 6105, 6127, 6151, 6179; Office_Menu.frm: 2838; RemoteVideo.frm: 657, 658, 664; Reservations.frm: 2207, 2239; Select_What.frm: 799; Server_Report_Menu.frm: 3871, 3872;
Subs.bas: 814, 897, 970, 971, 995, 1005, 1008, 1040, 1043, 1864, 1865, 1869, 1870, 2245, 2300, 2396, 2425, 2431, 4379, 4383, 4387, 4391, 4395, 4399, 4403, 4407, 4411, 4415, 4419, 4423, 4425, 4507, 4518, 10759, 10760, 10766, 10767, 10771, 10772, 10889, 10890, 10950, 10951, 11002, 11003, 12767, 21977, 24526, 24546, 24766, 27718, 27731, 29220, 29222, 33419, 33420, 33441, 33442, 33467, 33468, 33701, 33702, 33716, 33717, 33763, 33764;
Time_Reports.frm: 5497; functions.bas: 12901, 14055, 14056, 15963;
AuditDeposit.frm: 462 (in a trailing comment — verify before editing).

**All remaining `cdate([Col])` (no `#`) sites (file:line):**
AuditDeposit.frm: 434, 481, 483; Closing_Report.frm: 3176;
Display.frm: 2421, 2423, 2440, 2568, 2615, 2817, 4587, 4674, 4688, 4689, 4690, 4720, 4721, 4722, 5210, 6764, 6981, 6984, 6989, 6992, 7139, 7187, 7366, 7416, 7455, 7491, 7525, 7563, 7599, 7658, 7708, 7812, 7854, 7977, 7982, 7993, 7998, 8009, 8014, 8025, 8030, 8041, 8046, 8062, 8067, 8078, 8083;
Financal_Reports.frm: 9120, 10635, 12176, 12319; Mix.frm: 5858; Office_Menu.frm: 2821; Subs.bas: 671, 10891, 27732, 33446.

---

## Construct B — `Val([Col] & '')` as a SQL function (3)

`Val([Terminated] & '') = 0` → `ISNULL(TRY_CONVERT(int,[Terminated]),0) = 0`

| file:line | original | rewrite |
|---|---|---|
| Cashier_Audit.frm:2260 | `Where Val([Terminated] & '') = 0 order by [Last Name] asc` | `Where ISNULL(TRY_CONVERT(int,[Terminated]),0) = 0 order by [Last Name] asc` |
| Office_Menu.frm:3083 | same | same |
| Red_Book.frm:3776 | same | same |

---

## Construct C — `Format([Col],'HH:MM ampm')` in SQL (2, both Display.frm)

| file:line | original | rewrite |
|---|---|---|
| Display.frm:4691 | `and format([Hour],'HH:MM ampm') = '" & Format(RS![Hour] & "","HH:MM ampm") & "'"` | `and TRY_CAST([Hour] AS time) = TRY_CAST('" & Format(RS![Hour] & "","HH:nn") & "' AS time)"` |
| Display.frm:4723 | same | same |

---

## Construct D — `cast(... as date/datetime/int)` → `TRY_CAST` / `TRY_CONVERT` (309)

`cast([Col] as date)` → `TRY_CAST([Col] AS date)`. For text columns cast to int use
`TRY_CONVERT(int, [Col])`. For `cast(...) + cast(...)` datetime arithmetic, wrap the
sum in parentheses and ISO-format the VB-side bound.

**Representative rewrites:**

| file:line | original | rewrite |
|---|---|---|
| AccountTip.frm:761 | `Order By cast([Date] as date) DESC` | `Order By TRY_CAST([Date] AS date) DESC` |
| Assign_Name.frm:7809 | `cast([Date] as datetime) between '" & a & "' and '" & b & "'"` | `TRY_CAST([Date] AS datetime) between " & SqlDateTime(a) & " and " & SqlDateTime(b)` |
| Assign_Name.frm:8179 | `Cast([date] as datetime) Between '01/01/2000' and '" & a & "'"` | `TRY_CAST([date] AS datetime) Between '2000-01-01' and " & SqlDateTime(a)` |
| Closing_Report.frm:4365 | `Select Distinct cast([Order Date] as datetime) from [Closed Checks]` | `Select Distinct TRY_CAST([Order Date] AS datetime) from [Closed Checks]` |
| Display.frm:3202 | `from [AuditSales] Where cast([Date] as date) = '" & StartDate & "'"` | `from [AuditSales] Where TRY_CAST([Date] AS date) = '" & StartDate & "'"` |
| Display.frm:4555 | `Where cast([Start Date] as datetime) + cast([start time] as datetime) ...` | `Where (TRY_CAST([Start Date] AS datetime) + TRY_CAST([start time] AS datetime)) ...` |

**`cast(... as int)` (text column → use TRY_CONVERT):**
- RequestForm.frm:4959 — `cast([Check Number] / 1000 as int)` → `TRY_CONVERT(int, [Check Number] / 1000)`
- Server_Report_Menu.frm:3547, 3548 — `cast([Check Number] as int)` → `TRY_CONVERT(int, [Check Number])`

**All remaining D sites (file:line):**
Display.frm: 2239, 2944, 3209, 3356, 3375, 3814, 3877, 4007, 4577, 4585, 5093, 5095, 5795, 5797, 6689;
DriverMenu.frm: 968, 979, 995, 1098, 1117, 1134;
Financal_Reports.frm: 9718, 10092, 10204, 10629, 10631, 11878, 12017, 12162, 12163, 12172, 12173;
GFunctions.cls: 26; Office_Menu.frm: 2505, 2517, 2564, 3105;
PhoneOrders.frm: 9402, 9527, 9903, 9905, 10096, 10150, 10272, 10274, 10375, 11074, 11085, 11087, 11100, 11101, 11378, 11380, 11418, 11420, 11430, 11432, 11447, 11449;
Red_Book.frm: 3796;
Reservations.frm: 1708, 1848, 1849, 1856, 1887, 1891, 1911, 1938, 1939, 1951, 2010, 2209, 2241, 2454, 2590, 2591;
Select_What.frm: 634, 892, 893;
Server_Report_Menu.frm: 2947, 3106, 3310, 3382, 3398, 3706, 3713, 3720, 3895, 3899, 4011;
Subs.bas: 463, 695, 736, 738, 1599, 1774, 1963, 1967, 1971, 1975, 1979, 1985, 1989, 1993, 2052, 2056, 2061, 2065, 2234, 4433, 4437, 4441, 4443, 4450, 4452, 4454, 4457, 4459, 4462, 4464, 4466, 4468, 4470, 6730, 7694, 7697, 7700, 7703, 7727, 7775, 7959, 8096, 8126, 8150, 8319, 8321, 8502, 8726, 8737, 8744, 8746, 8748, 8787, 8789, 8791, 9235, 9282, 14352, 14400, 14418, 14420, 16881, 16889, 17216, 17217, 17473, 17487, 17491, 17594, 17657, 18687, 19025, 19086, 19117, 19190, 19235, 19322, 19389, 19419, 21295, 21325, 21330, 21368, 21455, 21459, 21463, 21506, 24937, 25438, 25675, 25678, 25745, 25748, 25810, 27685, 30051, 32719, 32963, 32994, 33818;
Time_Reports.frm: 2963, 2964, 3617, 3621, 3777, 3779, 4077, 4124, 4231, 4253, 4279, 4288, 4304, 4321, 4322, 4399, 4401, 4415, 4417, 5100, 5118, 5119, 5160, 5161, 5216, 5332, 5335, 5558, 5621, 5623, 5785, 5890;
Time_Sheet.frm: 3682, 3683, 3803, 3921, 4142, 4318, 4430, 4517, 4521, 4523, 4616, 4621, 4895, 4928, 5020, 5069, 5075, 5439;
WaitList.frm: 2167, 2174, 2181, 2364, 2371, 2376, 2411, 2432, 2433, 2551, 2624, 2629, 2662, 2669;
functions.bas: 576, 577, 611, 3578, 3616, 4122, 4336, 4360, 4385, 5742, 5972, 5984, 5997, 7041, 7050, 7135, 7421, 7442, 9193, 9195, 9523, 9535, 9634, 9736, 9770, 9810, 9855, 9888, 12923, 12934, 13194, 13258, 13260, 13309, 13824.

---

## Construct E — `&` between SQL columns/literals (3)

The only in-SQL `&` operator is `[Terminated] & ''` inside the Construct-B lines
(Cashier_Audit.frm:2260, Office_Menu.frm:3083, Red_Book.frm:3776). The B rewrite to
`ISNULL(TRY_CONVERT(int,[Terminated]),0)` removes the `& ''` entirely. No other
column-to-column/literal `&` exists in the live folder (all other `&` is VB-side
string building — leave as-is).

---

## Construct F — `Nz(` / `IIf(` in SQL: none

All `Nz(`/`IIf(` operate on VB recordset fields/variables, not SQL text. Nothing to
change. (If any are added later: `Nz(x,y)`→`ISNULL(x,y)`, `IIf(a,b,c)`→`CASE WHEN a
THEN b ELSE c END`.)

---

## Cautions

- **AuditDeposit.frm:462** matched the `#` search but the date fragment is in a
  trailing comment; the live SQL is just `"Select * from [Audits]"`. Verify, likely no change.
- Keep VB-side `CDate(...)` only where it normalizes input before `Format(...,"yyyy-mm-dd")`.
  The correctness key is the **ISO `yyyy-mm-dd[ hh:nn:ss]`** output.
- Test each rewritten report against a date range that includes month days > 12 — that's
  where ambiguous `MM/DD` vs `DD/MM` bugs surface.
