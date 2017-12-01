using System.Linq;
using ClosedXML.Report.Tests.TestModels;
using FluentAssertions;
using Xunit;

namespace ClosedXML.Report.Tests
{
    public class DbContextTests
    {
        [Fact]
        public void Customers_should_not_be_empty()
        {
            using (var db = new DbDemos())
            {
                db.customers.ToList().Should().NotBeEmpty();
            }
        }

        [Fact]
        public void Orders_should_not_be_empty()
        {
            using (var db = new DbDemos())
            {
                db.orders.ToList().Should().NotBeEmpty();
            }
        }
    }
}