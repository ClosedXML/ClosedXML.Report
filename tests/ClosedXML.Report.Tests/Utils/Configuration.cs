using System;
using Microsoft.Extensions.Configuration;

namespace ClosedXML.Report.Tests
{
    public class Configuration
    {
        public IConfigurationRoot Config { get; set; }
        public Configuration()
        {
            Config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", false, true)
                .Build();
            
        }

        public string DefaultConnectionString => Config.GetConnectionString("Default");
    }
}
