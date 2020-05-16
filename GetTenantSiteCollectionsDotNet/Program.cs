using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;

namespace GetTenantSiteCollectionsDotNet
{
    class Program
    {
        static void Main(string[] args)
        {
            string userName = GetFromConfigOrUser("username", "User");
            var password = GetSPOSecureStringPassword(GetFromConfigOrUser("password", "Password"));
            var adminUrl = GetFromConfigOrUser("adminurl", "SharePoint Admin URL");
            var outFilePath = $"{addTrailingbackslash(GetFromConfigOrUser("outputpath", "Output Path"))}{RemoveTrailingSlash(adminUrl).ToLower().Replace("-admin.sharepoint.com", "").Replace("https://", "")}_{DateTime.Now.ToString("yyyy-MM-dd_HHmmss")}.csv";

            try
            {
                using (ClientContext tenantContext = new ClientContext(adminUrl))
                {
                    //Authenticating with Tenant Admin
                    tenantContext.Credentials = new SharePointOnlineCredentials(userName, password);
                    WriteCSV((new Tenant(tenantContext)).GetSiteCollections(includeOD4BSites: true).ToList(), outFilePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Message: " + ex.Message);
                Console.ReadKey();
            }


            Console.WriteLine("press any key to continue...");
            Console.ReadKey();

        }

        private static string addTrailingbackslash(string v)
        {
            return v + (!v.EndsWith("\\") ? "\\" : "");
        }

        private static string GetFromConfigOrUser(string appSetting, string prompt)
        {
            var settingVal = ConfigurationManager.AppSettings[appSetting];
            if (settingVal == null || settingVal.Length == 0)
            {
                Console.Write($"{prompt}: ");
                settingVal = Console.ReadLine();
            }
            return settingVal;
        }

        private static string RemoveTrailingSlash(string adminUrl)
        {
            int lastSlash = adminUrl.LastIndexOf('/');
            adminUrl = (lastSlash > -1) ? adminUrl.Substring(0, lastSlash) : adminUrl;
            return adminUrl;
        }

        private static SecureString GetSPOSecureStringPassword(string password)
        {
            try
            {
                var secureString = new SecureString();
                foreach (char c in password)
                {
                    secureString.AppendChar(c);
                }
                return secureString;
            }

            catch
            {
                throw;
            }
        }

        public static void WriteCSV<T>(IEnumerable<T> items, string path)
        {
            Type itemType = typeof(T);
            var props = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                .OrderBy(p => p.Name);

            using (var writer = new StreamWriter(path))
            {
                writer.WriteLine(string.Join(", ", props.Select(p => p.Name)));

                foreach (var item in items)
                {
                    writer.WriteLine(string.Join(", ", props.Select(p => p.GetValue(item, null))));
                }
            }
        }
    }

}
