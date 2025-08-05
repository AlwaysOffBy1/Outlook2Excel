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
        public static bool IsOnErrorSendEmail { get; private set; } = false;
        public static string OnErrorSendEmailSMTPPath { get; private set; } = "";
        public static string OnErrorSendEmailFrom { get; private set; } = "";
        public static string OnErrorSendEmailFromName { get; private set; } = "Unnamed Outlok2Excel Service";
        public static string OnErrorSendEmailSubject { get; private set; } = "No Subject";
        public static string[] OnErroSendEmailTo { get; private set; } = new string[0];
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
            PrimaryKey = config["PrimaryKey"] ?? PrimaryKey;
            IsContainsPrimaryKey = PrimaryKey != "";
            ExcelFilePath = config["ExcelFilePath"] ?? ExcelFilePath;
            DaysToGoBack = TryConvertToInt(config["DaysToGoBack"]) ?? DaysToGoBack;
            RegexMap = ImportEmailMappings(config);
            TimerInterval = TryConvertToInt(config["TimerInterval"]) ?? TimerInterval;
            SubjectFilter = config["SubjectFilter"] ?? SubjectFilter;
            FromFilter = config["FromFilter"] ?? FromFilter;
            OrganizeBy = config["OrganizeBy"] ?? "EmailDate";

            IsOnErrorSendEmail = TryConvertToBool(config["IsOnErrorSendEmail"]) ?? IsOnErrorSendEmail;
            OnErrorSendEmailSMTPPath = config["OnErrorSendEmailSMTPPath"] ?? OnErrorSendEmailSMTPPath;
            OnErrorSendEmailFrom = config["OnErrorSendEmailFrom"] ?? "";
            OnErrorSendEmailFromName = config["OnErrorSendEmailFromName"] ?? OnErrorSendEmailFromName;
            OnErrorSendEmailSubject = config["OnErrorSendEmailSubject"] ?? OnErrorSendEmailSubject;
            OnErroSendEmailTo = string.IsNullOrEmpty(config["OnErroSendEmailTo"]) ? Array.Empty<string>() : config["OnErroSendEmailTo"].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

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
        private static bool? TryConvertToBool(string? value)
        {
            return bool.TryParse(value, out bool result) ? result : null;
        }
    }
}
