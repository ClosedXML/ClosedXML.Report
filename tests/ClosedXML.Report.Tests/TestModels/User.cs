using System;
using System.Collections.Generic;
using System.Linq;
using Bogus;
using Bogus.DataSets;
using Bogus.Extensions;

namespace ClosedXML.Report.Tests.TestModels
{
    public enum Gender
    {
        Male,
        Female
    }

    public class User
    {
        public static List<User> Generate(int count)
        {
            var fruit = new[] {"apple", "banana", "orange", "strawberry", "kiwi"};

            var orderIds = 0;
            var testOrders = new Faker<SimpleOrder>()
                .RuleFor(o => o.OrderId, f => orderIds++)
                .RuleFor(o => o.Item, f => f.PickRandom(fruit))
                .RuleFor(o => o.Quantity, f => f.Random.Number(1, 10))
                .RuleFor(o => o.OrderDate, f => f.Date.Past())
                .RuleFor(o => o.ShipDate, f => f.Date.Recent(40))
                //A nullable int? with 80% probability of being null.
                .RuleFor(o => o.LotNumber, f => f.Random.Int(0, 100).OrNull(f, .8f));


            var rnd = new Randomizer();
            var userIds = 0;
            var testUsers = new Faker<User>()
                .CustomInstantiator(f => new User(userIds++, f.Random.Replace("###-##-####")))

                .RuleFor(u => u.Gender, f => f.PickRandom<Gender>())

                .RuleFor(u => u.FirstName, (f, u) => f.Name.FirstName((Name.Gender?) u.Gender))
                .RuleFor(u => u.LastName, (f, u) => f.Name.LastName((Name.Gender?) u.Gender))
                .RuleFor(u => u.Avatar, f => f.Internet.Avatar())
                .RuleFor(u => u.UserName, (f, u) => f.Internet.UserName(u.FirstName, u.LastName))
                .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
                .RuleFor(u => u.SomethingUnique, f => $"Value {f.UniqueIndex}")

                .RuleFor(u => u.CartId, f => Guid.NewGuid())
                .RuleFor(u => u.FullName, (f, u) => u.FirstName + " " + u.LastName)
                .RuleFor(u => u.Orders, f => testOrders.Generate(rnd.Int(1, 5)).ToList())
                .FinishWith((f, u) => { Console.WriteLine("User Created! Id={0}", u.Id); });

            return testUsers.Generate(count);
        }

        public User(int id, string phoneNumber)
        {
            Id = id;
            PhoneNumber = phoneNumber;
        }

        public int Id { get; }
        public string PhoneNumber { get; }
        public Gender Gender { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Avatar { get; set; }
        public string UserName { get; set; }
        public string Email { get; set; }
        public string SomethingUnique { get; set; }
        public Guid CartId { get; set; }
        public string FullName { get; set; }
        public List<SimpleOrder> Orders { get; set; }
    }

    public class SimpleOrder
    {
        public int OrderId { get; set; }
        public string Item { get; set; }
        public int Quantity { get; set; }
        public int? LotNumber { get; set; }
        public DateTime OrderDate { get; set; }
        public DateTime ShipDate { get; set; }
    }
}
