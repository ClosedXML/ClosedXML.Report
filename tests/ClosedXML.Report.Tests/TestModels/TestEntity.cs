using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Report.Tests.TestModels
{
    public class TestEntity
    {
        public string Name { get; set; }
        public string Role { get; set; }
        public int Age { get; set; }
        public int[] Hours { get; set; }
        public Address Address { get; set; }

        public TestEntity(string name, string role, int age, int[] hours)
        {
            Hours = hours;
            Name = name;
            Role = role;
            Age = age;
        }

        public static IEnumerable<TestEntity> GetTestData(int rowCount)
        {
            return new[]
            {
                new TestEntity("John Smith", "Developer", 24, new [] { 6, 8, 4 }) {Address = new Address("USA", "NY", "94, Reade St")},
                new TestEntity("James Smith", "Analyst", 37, new [] { 3, 5, 7 }) {Address = new Address("USA", "Dallas", "5, Ross ave")},
                new TestEntity("Jim Smith", "Manager", 31, new[] { 2, 9, 1 }) {Address = new Address("USA", "Miami", "16, Indian Creek Dr")},
                new TestEntity("Chuck Norris", "Actor", 76, new [] { 7, 14, 2 }) {Address = new Address("USA", "Oklahoma", "9, Reade Rd")},
                new TestEntity("Dirk Benedict", "Actor", 71, new [] { 4, 9, 1 }) {Address = new Address("USA", "Montana", "7, Ross St, Helena")},
                new TestEntity("Kenneth Lauren Burns", "Producer", 63, new[] { 9, 1, 2 }) {Address = new Address("USA", "NY", "13, Indian Creek Dr, Brooklyn")},
            }.Take(rowCount);
        }
    }
}