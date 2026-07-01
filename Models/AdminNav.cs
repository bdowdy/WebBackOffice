namespace MSaleWebServer.Models;

public record AdminNavLink(string Title, string Href, string Sub);

public record AdminNavSection(string Id, string Name, string Accent, string Icon, AdminNavLink[] Links);

/// <summary>
/// Single source of truth for back-office navigation. Used by the shared
/// layout sidebar and the dashboard launcher (Views/MSaleAdmin/Index.cshtml).
/// Icon strings are heroicon (outline) path data rendered inline.
/// </summary>
public static class AdminNav
{
    public static readonly AdminNavSection[] Sections =
    [
        new(
            "dashboard", "MSale Dashboard", "gold",
            "M3 13.125C3 12.504 3.504 12 4.125 12h2.25c.621 0 1.125.504 1.125 1.125v6.75C7.5 20.496 6.996 21 6.375 21h-2.25A1.125 1.125 0 0 1 3 19.875v-6.75ZM9.75 8.625c0-.621.504-1.125 1.125-1.125h2.25c.621 0 1.125.504 1.125 1.125v11.25c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V8.625ZM16.5 4.125c0-.621.504-1.125 1.125-1.125h2.25C20.496 3 21 3.504 21 4.125v15.75c0 .621-.504 1.125-1.125 1.125h-2.25a1.125 1.125 0 0 1-1.125-1.125V4.125Z",
            [
                new("Sales Dashboard", "https://app.microsaledashboard.net/auth/login", "Open the MSale live sales dashboard — view real-time sales, labor, and reporting."),
            ]),
        new(
            "discounts", "Discounts", "gold",
            "M9.568 3H5.25A2.25 2.25 0 0 0 3 5.25v4.318c0 .597.237 1.17.659 1.591l9.581 9.581c.699.699 1.78.872 2.607.33a18.095 18.095 0 0 0 5.223-5.223c.542-.827.369-1.908-.33-2.607L11.16 3.66A2.25 2.25 0 0 0 9.568 3Z M6 6h.008v.008H6V6Z",
            [
                new("Discount Events", "/MSaleAdmin/DiscountEventAssignments", "Assign discounts to delivery, gratuity, service charge, loyalty, and other events."),
                new("Discount and Coupon Assignments", "/MSaleAdmin/Discounts", "Assign discounts to menu items, modifiers, and departments."),
                new("Discount Groups", "/MSaleAdmin/CouponAssignments", "Manage discount grouping and discount setup lists."),
            ]),
        new(
            "taxes", "Taxes", "gold",
            "M2.25 18.75a60.07 60.07 0 0 1 15.797 2.101c.727.198 1.453-.342 1.453-1.096V18.75M3.75 4.5v.75A.75.75 0 0 1 3 6h-.75m0 0v-.375c0-.621.504-1.125 1.125-1.125H20.25M2.25 6v9m18-10.5v.75c0 .414.336.75.75.75h.75m-1.5-1.5h.375c.621 0 1.125.504 1.125 1.125v9.75c0 .621-.504 1.125-1.125 1.125h-.375m1.5-1.5H21a.75.75 0 0 0-.75.75v.75m0 0H3.75m0 0h-.375a1.125 1.125 0 0 1-1.125-1.125V15m1.5 1.5v-.75A.75.75 0 0 0 3 15h-.75M15 10.5a3 3 0 1 1-6 0 3 3 0 0 1 6 0Zm3 0h.008v.008H18V10.5Zm-12 0h.008v.008H6V10.5Z",
            [
                new("Sales Tax", "/MSaleAdmin/TaxSetup", "Manage tax names, tax rates, and tax tables."),
            ]),
        new(
            "menu-pricing", "Menu & Pricing", "gold",
            "M3.75 6A2.25 2.25 0 0 1 6 3.75h2.25A2.25 2.25 0 0 1 10.5 6v2.25a2.25 2.25 0 0 1-2.25 2.25H6a2.25 2.25 0 0 1-2.25-2.25V6ZM3.75 15.75A2.25 2.25 0 0 1 6 13.5h2.25a2.25 2.25 0 0 1 2.25 2.25V18a2.25 2.25 0 0 1-2.25 2.25H6A2.25 2.25 0 0 1 3.75 18v-2.25ZM13.5 6a2.25 2.25 0 0 1 2.25-2.25H18A2.25 2.25 0 0 1 20.25 6v2.25A2.25 2.25 0 0 1 18 10.5h-2.25a2.25 2.25 0 0 1-2.25-2.25V6ZM13.5 15.75a2.25 2.25 0 0 1 2.25-2.25H18a2.25 2.25 0 0 1 2.25 2.25V18A2.25 2.25 0 0 1 18 20.25h-2.25A2.25 2.25 0 0 1 13.5 18v-2.25Z",
            [
                new("Timed Pricing", "/MSaleAdmin/TimedPricing", "Manage happy hour, day based pricing, and timed price changes."),
                new("Order Buttons", "/MSaleAdmin/OrderButtons", "Program order buttons, page numbers, captions, and positions."),
                new("Button Layout", "/MSaleAdmin/ButtonLayout", "Manage menu button layout, colors, sizes, icons, and placement."),
                new("Department Setup", "/MSaleAdmin/Reportcategories", "Manage departments, colors, sizes, icons, and placement."),
                new("Category Setup", "/MSaleAdmin/CategorySetup", "Manage categories, colors, sizes, icons, and placement."),
                new("Meal Period Setup", "/MSaleAdmin/MealPeriods", "Manage meal periods, colors, sizes, icons, and placement."),
                new("Modifier Prefix Setup", "/MSaleAdmin/ModifierPrefixes", "Manage modifier prefixes, colors, sizes, icons, and placement."),
                new("SubCategory Setup", "/MSaleAdmin/SubCategories", "Manage sub categories, colors, sizes, icons, and placement."),
                new("Shift Level Setup", "/ShiftLevel", "Create and manage prefix/suffix shift keys — assign price levels, categories, timed event overrides, and PLU transformer codes."),
            ]),
        new(
            "security", "Security & Access", "gold",
            "M9 12.75 11.25 15 15 9.75m-3-7.036A11.959 11.959 0 0 1 3.598 6 11.99 11.99 0 0 0 3 9.749c0 5.592 3.824 10.29 9 11.623 5.176-1.332 9-6.03 9-11.622 0-1.31-.21-2.571-.598-3.751h-.152c-3.196 0-6.1-1.248-8.25-3.285Z",
            [
                new("Access Levels", "/MSaleAdmin/AccessLevels", "Manage employee access level names and security numbers."),
                new("Menu Access", "/MSaleAdmin/MenuAccess", "Assign access requirements to menus and programming areas."),
                new("Job Titles", "/MSaleAdmin/JobTitles", "Manage job titles, tipped positions, tip pool percent, 7shifts job codes, and drawer access."),
                new("Employee Maintenance", "/MSaleAdmin/Employees", "Manage employee names and security numbers."),
                new("Check Review", "/CheckReview", "Review closed checks and payments."),
            ]),
        new(
            "store-setup", "Store Setup", "gold",
            "M13.5 21v-7.5a.75.75 0 0 1 .75-.75h3a.75.75 0 0 1 .75.75V21m-4.5 0H2.36m11.14 0H18m0 0h3.64m-1.39 0V9.349m-16.5 11.65V9.35m0 0a3.001 3.001 0 0 0 3.75-.615A2.993 2.993 0 0 0 9.75 9.75c.896 0 1.7-.393 2.25-1.016a2.993 2.993 0 0 0 2.25 1.016c.896 0 1.7-.393 2.25-1.015a3.001 3.001 0 0 0 3.75.614m-16.5 0a3.004 3.004 0 0 1-.621-4.72l1.189-1.19A1.5 1.5 0 0 1 5.378 3h13.243a1.5 1.5 0 0 1 1.06.44l1.19 1.189a3 3 0 0 1-.621 4.72m-13.5 8.65h3.75a.75.75 0 0 0 .75-.75V13.5a.75.75 0 0 0-.75-.75H6.75a.75.75 0 0 0-.75.75v3.75c0 .415.336.75.75.75Z",
            [
                new("Tenders", "/MSaleAdmin/TenderAdmin", "Manage tender names."),
                new("Account Maintenance", "/MSaleAdmin/AssignName", "Manage house and charge accounts, balances, limits, and status."),
                new("Account Activity Report", "/MSaleAdmin/AccountActivity", "Print house account charges and payments over a date range with a running balance."),
                new("Suggested Tip", "/MSaleAdmin/SuggestedTips", "Tip helper setup."),
                new("Tare", "/MSaleAdmin/TareSetup", "Manage tare names and tare weights."),
                new("Refund Reasons", "/MSaleAdmin/RefundReasons", "Add and edit refund reasons used by the POS."),
                new("Preferences", "/MSaleAdmin/Preferences", "Manage preference names used in store programming."),
                new("Store Paid In", "/MSaleAdmin/StorePaidIn", "Add, edit, and delete paid-in reason names."),
                new("Void Reasons", "/MSaleAdmin/VoidReasons", "Manage void reasons and inventory removal options."),
                new("Special Requests", "/MSaleAdmin/SpecialRequests", "Manage special request names, instructions, taxes, remotes, and open item options."),
                new("Nutritional Setup", "/MSaleAdmin/Nutritional", "Manage nutritional calorie information."),
                new("Non Sale Revenue", "/MSaleAdmin/RevenueAssignments", "Manage paid items that are not normal sales revenue."),
                new("Master Departments", "/MSaleAdmin/MasterDepartments", "Manage master department names used for reporting."),
                new("Receipt Messages", "/MSaleAdmin/ReceiptMessages", "Manage receipt footer and customer messages."),
                new("Daily Specials", "/MSaleAdmin/DailySpecials", "Manage daily specials and flash messages."),
                new("Table Layout Editor", "/TableLayout", "Design dining room floor plans — add, position, and size tables on a drag-and-drop canvas with multiple rooms and floor backgrounds."),
                new("Language Editor", "/MSaleAdmin/LanguageEditor", "Maintain POS text translations — edit the English phrase and its translation used for on-screen language translation."),
            ]),
        new(
            "printers", "Printers & Hardware", "gold",
            "M6.72 13.829c-.24.03-.48.062-.72.096m.72-.096a42.415 42.415 0 0 1 10.56 0m-10.56 0L6.34 18m10.94-4.171c.24.03.48.062.72.096m-.72-.096L17.66 18m0 0 .229 2.523a1.125 1.125 0 0 1-1.12 1.227H7.231c-.662 0-1.18-.568-1.12-1.227L6.34 18m11.318 0h1.091A2.25 2.25 0 0 0 21 15.75V9.456c0-1.081-.768-2.015-1.837-2.175a48.055 48.055 0 0 0-1.913-.247M6.34 18H5.25A2.25 2.25 0 0 1 3 15.75V9.456c0-1.081.768-2.015 1.837-2.175a48.041 48.041 0 0 1 1.913-.247m10.5 0a48.536 48.536 0 0 0-10.5 0m10.5 0V3.375c0-.621-.504-1.125-1.125-1.125h-8.25c-.621 0-1.125.504-1.125 1.125v3.659M18 10.5h.008v.008H18V10.5Zm-12 0h.008v.008H6V10.5Z",
            [
                new("Printer Setup", "/MSaleAdmin/PrinterSetup", "Configure receipt, remote, and logo printers, cash drawers, receipt routing, and Hibachi Grill station."),
                new("Receipt Print Size", "/MSaleAdmin/ReceiptPrintSize", "Set large or small print for menu item and modifier lines on the receipt, per printer."),
            ]),
        new(
            "frequent-diner", "Frequent Diner Loyalty", "gold",
            "M18 18.72a9.094 9.094 0 0 0 3.741-.479 3 3 0 0 0-4.682-2.72m.94 3.198.001.031c0 .225-.012.447-.037.666A11.944 11.944 0 0 1 12 21c-2.17 0-4.207-.576-5.963-1.584A6.062 6.062 0 0 1 6 18.719m12 0a5.971 5.971 0 0 0-.941-3.197m0 0A5.995 5.995 0 0 0 12 12.75a5.995 5.995 0 0 0-5.058 2.772m0 0a3 3 0 0 0-4.681 2.72 8.986 8.986 0 0 0 3.74.477m.94-3.197a5.971 5.971 0 0 0-.94 3.197M15 6.75a3 3 0 1 1-6 0 3 3 0 0 1 6 0Zm6 3a2.25 2.25 0 1 1-4.5 0 2.25 2.25 0 0 1 4.5 0Zm-13.5 0a2.25 2.25 0 1 1-4.5 0 2.25 2.25 0 0 1 4.5 0Z",
            [
                new("Frequent Diner Setup", "/FrequentDiner", "Manage loyalty customers, plans, qualifying items, timed bonus events, and view customer reward progress."),
            ]),
        new(
            "timekeeping", "Time Keeping & Reporting", "gold",
            "M12 6v6h4.5m4.5 0a9 9 0 1 1-18 0 9 9 0 0 1 18 0Z",
            [
                new("Time Records Report",    "/TimeRecordsReport",          "Pre-print closed time records by employee for a selected date range."),
                new("Payroll Summary Report", "/MSaleAdmin/PayrollSummary",  "Employee payroll summary reporting."),
                new("Audit Report",           "/AuditReport",                "Register and operator close-out audit by date range — sales, discounts, voids, cash, deposits, over/short, and tips."),
                new("Server Income Report",   "/ServerIncomeReport",         "Server tip and charge-sales income by date range — cash tips, charge tips, tip share, total tips, and charge tip %."),
                new("Labor Reports Hub",      "/TimeReports",                "Pay period labor, overtime watch, hourly labor, punch report, breaks, hours by department, avg rate exception, individual daily, and more."),
                new("Who's On the Clock",     "/TimeReports",                "Live view of employees currently clocked in."),
                new("Phone List",             "/TimeReports",                "Print active employee phone numbers."),
                new("Access Codes",           "/TimeReports",                "Print employee access levels and codes."),
                new("Time Record Editor",     "/MSaleAdmin/TimeRecordEditor","Edit employee punches, breaks, tips, sales, and pay rates."),
                new("Tip Editor",             "/MSaleAdmin/TipEditor",       "Edit cash tips, credit card tips, tips paid out, tip share, and gratuity."),
                new("Time Clock Messages",    "/MSaleAdmin/TimeMessages",    "Send time clock messages to employees or departments."),
            ]),
        new(
            "modifiers", "Modifiers & Menu Items", "gold",
            "M10.5 6h9.75M10.5 6a1.5 1.5 0 1 1-3 0m3 0a1.5 1.5 0 1 0-3 0M3.75 6H7.5m3 12h9.75m-9.75 0a1.5 1.5 0 0 1-3 0m3 0a1.5 1.5 0 0 0-3 0m-3.75 0H7.5m9-6h3.75m-3.75 0a1.5 1.5 0 0 1-3 0m3 0a1.5 1.5 0 0 0-3 0m-9.75 0h9.75",
            [
                new("Modifier Maintenance",  "/ModifierMaintenance",           "Create modifier names and manage modifier items."),
                new("Modifier Price Levels", "/MSaleAdmin/ModifierPriceLevels", "Maintain modifier price levels."),
                new("Modifier Details",      "/MSaleAdmin/ModifierDetails",     "Edit modifier item costs, taxes, printers, aliases, PLU, and options."),
                new("Menu Item Maintenance", "/MenuItem",                       "Add and edit menu items."),
                new("Menu List Report",      "/MenuListReport",                 "Printable list of all menu items by report category — PLU, price, tax, modifiers, and remotes."),
            ]),
        new(
            "front-of-house", "Front of House", "gold",
            "M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 0 1 2.25-2.25h13.5A2.25 2.25 0 0 1 21 7.5v11.25m-18 0A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75m-18 0v-7.5A2.25 2.25 0 0 1 5.25 9h13.5A2.25 2.25 0 0 1 21 11.25v7.5",
            [
                new("Reservations", "/Reservation", "View, add, and manage dining reservations — track party size, date, time, table assignment, and status."),
                new("Wait List",    "/WaitList",    "Live customer wait list — add walk-ins, track elapsed time, seat parties, and send table-ready notifications."),
            ]),
    ];
}
