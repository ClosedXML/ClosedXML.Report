using System.Collections.Generic;
using Bogus;

namespace ClosedXML.Report.Tests.TestModels
{
    public class TestOrder
    {
        public string OrderNumber { get; set; }
        public List<OrderItem> ProductsWithQuantities { get; set; } = new List<OrderItem>();

        public TestOrder()
        {
        }

        public TestOrder(string orderNumber)
        {
            OrderNumber = orderNumber;
        }

        public static IEnumerable<TestOrder> GetTestData(int rowCount)
        {
            var products = new[]
            {
                "Brioche", "Tart taten", "Apple pie", "Creamy croissants", "Toast with cream", "Scramble croissant",
                "Beunier donuts", "Profiteroles with Mascarpone", "Creme de parisien", "Chocolate Fondant", "Quiche with bacon", 
            };

            var orderItemFaker = new Faker<OrderItem>()
                .RuleFor(x => x.ProductName, f => f.PickRandom(products))
                .RuleFor(x => x.Quantity, f => f.Random.Number(1, 50));

            var orderFaker = new Faker<TestOrder>()
                .RuleFor(x => x.OrderNumber, f => f.Random.Number(100000, 999999).ToString())
                .RuleFor(x => x.ProductsWithQuantities, f => orderItemFaker.Generate(f.Random.Number(2, 10)));

            return orderFaker.Generate(rowCount);
        }
    }

    public class OrderItem
    {
        public OrderItem()
        {
        }

        public OrderItem(string productName, decimal quantity)
        {
            ProductName = productName;
            Quantity = quantity;
        }

        public string ProductName { get; set; }
        public decimal Quantity { get; set; }
    }

    public class TestEntity
    {
        public string Name { get; set; }
        public string Role { get; set; }
        public int Age { get; set; }
        public int[] Hours { get; set; }
        public Address Address { get; set; }

        public TestEntity()
        {
        }

        public TestEntity(string name, string role, int age, int[] hours)
        {
            Hours = hours;
            Name = name;
            Role = role;
            Age = age;
        }

        public static IEnumerable<TestEntity> GetTestData(int rowCount)
        {
            //var roles = new[] { "Developer", "Analyst", "Manager", "Actor", "Producer" };
            var addressFaker = new Faker<Address>()
                .RuleFor(o => o.Country, f => f.Address.Country())
                .RuleFor(o => o.City, f => f.Address.City())
                .RuleFor(o => o.Street, f => f.Address.StreetAddress());

            var testEntity = new Faker<TestEntity>()
                .StrictMode(true)
                .RuleFor(o => o.Name, f => f.Name.FullName())
                .RuleFor(o => o.Role, f => f.Name.JobTitle())
                .RuleFor(o => o.Age, f => f.Random.Number(20, 70))
                .RuleFor(o => o.Address, () => addressFaker)
                .RuleFor(o => o.Hours, f => new []{ f.Random.Number(2, 14), f.Random.Number(2, 14), f.Random.Number(2, 14) });

            return testEntity.Generate(rowCount);
        }
    }
}
