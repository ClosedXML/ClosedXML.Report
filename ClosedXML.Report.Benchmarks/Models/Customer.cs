namespace ClosedXML.Report.Benchmarks.Models;

public class Customer
{
    public required string Company { get; set; }

    public required string Addr1 { get; set; }

    public required string Addr2 { get; set; }


    public required string City { get; set; }

    public required string State { get; set; }

    public required string Country { get; set; }

    public required string Phone { get; set; }

    public required string Email { get; set; }

    public required string Zip { get; set; }

    public string Fax { get; set; } = string.Empty;

    public List<Order> Orders { get; init; } = [];
}