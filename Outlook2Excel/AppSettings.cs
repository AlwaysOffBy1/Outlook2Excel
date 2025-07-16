using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace Outlook2Excel
{
    static class AppSettings
    {
        public static string? Mailbox;
        public static string? PrimaryKey;
        public static string? ExcelFilePath;
        public static List<ImportObject>? ImportObjects;

        public static bool GetSettings()
        {
            //Get config from file
            var config = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettings.json")
                 .Build();
            if (config == null) return false;

            //Set vars
            Mailbox = config["Mailbox"];
            PrimaryKey = config["PrimaryKey"];
            ExcelFilePath = config["ExcelFilePath"];
            ImportObjects = ImportEmailMappings(config);

            //If any vars are null return false
            return Mailbox != null
                && PrimaryKey != null
                && ExcelFilePath != null
                && ImportObjects != null;
        }

        private static List<ImportObject>? ImportEmailMappings(IConfiguration config)
        {
            //Get list of keyvaluepairs to be imported, and import them as an "IConfigurationSection"
            var list = new List<ImportObject>();
            var section = config.GetSection("EmailMessageMapping");

            foreach (var child in section.GetChildren())
            {
                //If any key value pairs are blank or null, exit the program
                string? key = child.Key;
                string? value = child.Value;
                if (key == null || value == null) return null;

                list.Add(new ImportObject(key, value));
            }

            return list;
        }
    }
}
