using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace Outlook2Excel
{
    public static class AppSettings
    {
        public static string PrimaryKey { get; private set; } = "Subject";
        public static string ExcelFilePath { get; private set; } = "";
        public static Dictionary<string, string> RegexMap { get; set; } = new Dictionary<string, string>();
        public static bool IsContainsPrimaryKey { get; private set; } = false;
        public static int DaysToGoBack { get; set; } = 1;
        public static int TimerInterval { get; set; } = 5;
        public static string SubjectFilter { get; private set; } = "";
        public static string FromFilter { get; private set; } = "";
        public static string OrganizeBy { get; private set; } = "EmailDate";

        public static string FullFolderPath { get; private set; } = "";

        public static bool GetSettings()
        {
            //Get config from file
            var config = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettings.json")
                 .Build();
            if (config == null) return false;

            //Set vars
            FullFolderPath = config["FullFolderPath"] ?? "";
            PrimaryKey = config["PrimaryKey"] ?? "Subject";
            IsContainsPrimaryKey = PrimaryKey != "";
            ExcelFilePath = config["ExcelFilePath"] ?? "";
            DaysToGoBack = TryConvertToInt(config["DaysToGoBack"]) ?? DaysToGoBack;
            RegexMap = ImportEmailMappings(config);
            TimerInterval = TryConvertToInt(config["TimerInterval"]) ?? TimerInterval;
            SubjectFilter = config["SubjectFilter"] ?? string.Empty;
            FromFilter = config["FromFilter"] ?? string.Empty;
            OrganizeBy = config["OrganizeBy"] ?? "EmailDate";

            //If any mandatory vars are null return false
            return FullFolderPath != null
                && PrimaryKey != null
                && ExcelFilePath != null
                && RegexMap != null;
        }

        private static Dictionary<string,string> ImportEmailMappings(IConfiguration config)
        {
            //Get list of keyvaluepairs to be imported, and import them as an "IConfigurationSection"
            var list = new Dictionary<string, string>();
            var section = config.GetSection("EmailMessageMapping");

            foreach (var child in section.GetChildren())
            {
                //If any key value pairs are blank or null, exit the program
                string? key = child.Key;
                string? value = child.Value;
                if (key == null || value == null) return new Dictionary<string, string>();

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
