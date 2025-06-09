namespace ClosedXML.Report.Benchmarks.Models;

public class Order
{
    public required string OrderNo { get; set; }

    public DateTime SaleDate { get; set; }

    public DateTime ShipDate { get; set; }

    public string ShipToAddr1 { get; set; } = string.Empty;

    public string ShipToAddr2 { get; set; } = string.Empty;

    public string PaymentMethod { get; set; } = string.Empty;

    public int ItemsTotal { get; set; }

    public decimal TaxRate { get; set; }

    public decimal AmountPaid { get; set; }
}