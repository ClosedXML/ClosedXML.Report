using BenchmarkDotNet.Running;
using ClosedXML.Report.Benchmarks.Models;

namespace ClosedXML.Report.Benchmarks;

public class Program
{
    static void Main(string[] args)
    {
        BenchmarkRunner.Run(typeof(Program).Assembly);
    }
}

public class DataBuilder
{
    public Order CreateOrder(int orderNo = 1, int amount = 10000)
    {
        var orderNoStr = orderNo.ToString("D6");
        var isOdd = orderNo % 2 == 1;

        return new Order
        {
            AmountPaid = amount,
            ItemsTotal = isOdd ? 1 : 2,
            OrderNo = orderNoStr,
            PaymentMethod = isOdd ? "Credit" : "Visa",
            SaleDate = DateTime.Now,
            ShipDate = DateTime.Now,
            ShipToAddr1 = "ShipToAddr1",
            ShipToAddr2 = "ShipToAddr2",
            TaxRate = 0.1m
        };
    }

    public Customer Create()
    {
        var c = new Customer
        {
            City = "City",
            Addr1 = "1 Main St",
            Addr2 = "Townsville",
            Company = "Company Name",
            Country = "Australia",
            Email = "boss@gmail.com",
            Fax = "0011 123 456 789",
            Phone = "0011 123 456 789",
            State = "Qld",
            Zip = "4000",
            Orders = []
        };
        for (var i = 0; i < 10000; i++)
        {
            c.Orders.Add(CreateOrder(i, 100 + i));
        }

        return c;
    }
}