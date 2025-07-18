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
        public static Dictionary<string, string>? RegexMap;
        public static bool IsContainsPrimaryKey = false;
        public static int DaysToGoBack = 1;
        public static int TimerInterval = 5;

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
            IsContainsPrimaryKey = PrimaryKey != "";
            ExcelFilePath = config["ExcelFilePath"];
            DaysToGoBack = TryConvertToInt(config["DaysToGoBack"]) ?? DaysToGoBack;
            RegexMap = ImportEmailMappings(config);
            TimerInterval = TryConvertToInt(config["TimerInterval"]) ?? TimerInterval;

            //If any mandatory vars are null return false
            return Mailbox != null
                && PrimaryKey != null
                && ExcelFilePath != null
                && RegexMap != null;
        }

        private static Dictionary<string,string>? ImportEmailMappings(IConfiguration config)
        {
            //Get list of keyvaluepairs to be imported, and import them as an "IConfigurationSection"
            var list = new Dictionary<string, string>();
            var section = config.GetSection("EmailMessageMapping");

            foreach (var child in section.GetChildren())
            {
                //If any key value pairs are blank or null, exit the program
                string? key = child.Key;
                string? value = child.Value;
                if (key == null || value == null) return null;

                list.Add(key, value);
            }

            return list;
        }

        private static int? TryConvertToInt(string? value)
        {
            return int.TryParse(value, out int result) ? result : null;
        }
    }
}
