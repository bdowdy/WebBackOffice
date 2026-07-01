public class SystemPrinterDto
{
    public string PrinterName { get; set; } = "";
    public string PortName { get; set; } = "";
    public string DeviceName { get; set; } = "";
    public string FailedPrinter { get; set; } = "";
    public string GoOrderPrinter { get; set; } = "";
    public string DeliveryPrinter { get; set; } = "";
    public bool FirePrinter { get; set; }
}

public class CashDrawerDto
{
    public int DrawerNumber { get; set; }
    public string PortName { get; set; } = "";
    public string DataRate { get; set; } = "";
    public int Code1 { get; set; }
    public int Code2 { get; set; }
    public int Code3 { get; set; }
    public int Code4 { get; set; }
    public int Code5 { get; set; }
}