using System.Collections.Generic;
using System.Linq;
using LinqToDB;
using LinqToDB.Configuration;

namespace ClosedXML.Report.Tests
{
    public class ConnectionStringSettings : IConnectionStringSettings
    {
        public string ConnectionString { get; set; }
        public string Name { get; set; }
        public string ProviderName { get; set; }
        public bool IsGlobal => false;
    }

    public class DbSettings : ILinqToDBSettings
    {
        public IEnumerable<IDataProviderSettings> DataProviders => Enumerable.Empty<IDataProviderSettings>();
        public string DefaultConfiguration => "Default";
        public string DefaultDataProvider => ProviderName.SQLite;

        private readonly Configuration _configuration = new Configuration();

        public IEnumerable<IConnectionStringSettings> ConnectionStrings
        {
            get
            {
                yield return
                    new ConnectionStringSettings
                    {
                        Name = "Default",
                        ProviderName = ProviderName.SQLite,
                        ConnectionString = _configuration.DefaultConnectionString
                    };
            }
        }
    }
}
