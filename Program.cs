using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using IISInstallationTest.models;
using Microsoft.Win32;
using Microsoft.Web.Administration;

namespace IISInstallationTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Version version = GetIISVersion();
            Console.WriteLine("Major Version = " + version.Major.ToString());
            Console.WriteLine("Minor Version = " + version.Minor.ToString());

            if (version.Major >= 7)
            {
                Console.WriteLine("IIS 7 or higher found. Proceeding with installation...");
                InstallForSevenOrUp();
            }
            else
            {
                Console.WriteLine("IIS version is " + version.Major.ToString() + "." + version.Minor.ToString() + ". This does not meet the minimum requirements.");
                Console.WriteLine("Installation Aborted.");
            }
            Console.ReadLine();
        }


        public static Version GetIISVersion()
        {
            using (RegistryKey componentsKey = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\InetStp", false))
            {
                if (componentsKey == null)
                    return new Version(0, 0);

                int majorVersion = (int)componentsKey.GetValue("MajorVersion", -1);
                int minorVersion = (int)componentsKey.GetValue("MinorVersion", -1);

                if (majorVersion != -1 && minorVersion != -1)
                {
                    return new Version(majorVersion, minorVersion);
                }

                return new Version(0, 0);
            }
        }



        private static void InstallForSevenOrUp()
        {
            try
            {
                CreateApplicationPools();
                AddApplications();
                Console.WriteLine("Installation Completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Some error occured: " + ex.Message);
                Console.WriteLine("Installation aborted");
            }
        }


        private static void AddApplications()
        {
            const string windowsSectionName = "system.webServer/security/authentication/windowsAuthentication";
            const string anonymousSectionName = "system.webServer/security/authentication/anonymousAuthentication";
            bool anonymousUnlocked = UnlockSection(anonymousSectionName);
            bool windowsUnlocked = UnlockSection(windowsSectionName);

            using (var serverManager = ServerManager.OpenRemote("localhost"))
            {
                Console.WriteLine("Adding Applications...");
                Site site = GetSiteNumber(serverManager);


                try
                {
                    Application serviceApplication = site.Applications.Add("/Cevi/Services/WCF/Vergunningen.Service",
                        @"C:\inetpub\wwwroot\Cevi\Services\WCF\Vergunningen.Service");
                    serviceApplication.ApplicationPoolName = "VergunningenPool";
                    Console.WriteLine("Converted Vergunningen.Service to application and added it to VergunningenPool");
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine("Error Vergunningen.Service: " + ex.Message);
                }

                try
                {
                    Application reportingApplication = site.Applications.Add("/Cevi/Services/WCF/Vergunningen.Reporting",
                        @"C:\inetpub\wwwroot\Cevi\Services\WCF\Planregister.Service");
                    reportingApplication.ApplicationPoolName = "VergunningenPool";
                    Console.WriteLine("Converted Vergunningen.Reporting to application and added it to VergunningenPool");
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine("Error Vergunningen.Reporting: " + ex.Message);
                }

                try
                {
                    Application planregisterApplication = site.Applications.Add("/Cevi/Services/WCF/Planregister.Service",
                        @"C:\inetpub\wwwroot\Cevi\Services\WCF\Planregister.Service");
                    planregisterApplication.ApplicationPoolName = "PlanregisterPool";
                    Console.WriteLine("Converted Planregister.Service to application and added it to PlanregisterPool");
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine("Error Planregister.Service: " + ex.Message);
                }

                try
                {
                    Application dbaApplication = site.Applications.Add("/Cevi/Services/WCF/DBA.Service",
                        @"C:\inetpub\wwwroot\Cevi\Services\WCF\DBA.Service");
                    dbaApplication.ApplicationPoolName = "DBAImportPool";
                    Console.WriteLine("Converted DBA.Service to application and added it to DBAImportPool");
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine("Error DBA.Service: " + ex.Message);
                }

                try
                {
                    Application vergunningenApplication = site.Applications.Add("/Cevi/Vergunningen.Web",
                        @"C:\inetpub\wwwroot\Cevi\Vergunningen.Web");
                    vergunningenApplication.ApplicationPoolName = "VergunningenPool";
                    Console.WriteLine("Converted Vergunningen.Web to application and added it to VergunningenPool");
                    Configuration config = vergunningenApplication.GetWebConfiguration();
                    try
                    {
                        if (anonymousUnlocked)
                        {
                            ConfigurationSection anonymousAuthenticationSection =
                                config.GetSection(anonymousSectionName);
                            anonymousAuthenticationSection["enabled"] = false;
                            Console.WriteLine("Deactivated anonymous authentication");
                        }
                        else
                        {
                            Console.WriteLine("Unable to unlock anonymousSection");
                        }

                    }
                    catch (FileLoadException ex)
                    {
                        Console.WriteLine("Error Vergunningen.Web: " + ex.Message);
                    }
                    try
                    {

                        if (windowsUnlocked)
                        {
                            ConfigurationSection windowsAuthenticationSection =
                            config.GetSection(windowsSectionName);
                            windowsAuthenticationSection["enabled"] = true;
                            Console.WriteLine("Activated windows authentication");
                        }
                        else
                        {
                            Console.WriteLine("Unable to unlock windowsSection");
                        }

                    }
                    catch (FileLoadException ex)
                    {
                        Console.WriteLine("Error Vergunningen.Web: " + ex.Message);
                    }

                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine("Error Vergunningen.Web: " + ex.Message);
                }
                serverManager.CommitChanges();
            }

        }

        private static Site GetSiteNumber(ServerManager serverManager)
        {
            int siteId = 1;
            Console.WriteLine(
                "Geef het nummer in van de Site waarop de applicaties mogen komen. Dit kan je vinden door in IIS manager op de site Rechtsklik > geavanceerde Instellingen > Id");
            Console.WriteLine("Dit staat reeds standaard ingesteld op Default Site (id = 1).");
            string siteIdString = Console.ReadLine();
            if (string.IsNullOrEmpty(siteIdString))
            {
                Console.WriteLine("No number entered, will continue with Default Site with id 1.");
            }
            else
            {
                int testnr = 0;
                if (int.TryParse(siteIdString, out testnr) && testnr != 0)
                {
                    siteId = testnr;
                }
                else
                {
                    Console.WriteLine("Site is no number, try again");
                    GetSiteNumber(serverManager);
                }
            }

            Site site = null;
            try
            {
                site = serverManager.Sites.First(s => s.Id == siteId);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Site with id " + siteId + " not found, try again");
                GetSiteNumber(serverManager);
            }
            return site;
        }

        private static void CreateApplicationPools()
        {
            ImpersonatedUser user = GetImpersonatedUser();
            using (ServerManager serverManager = new ServerManager())
            {
                Console.WriteLine("Creating Application Pools...");

                try
                {
                    if (serverManager.ApplicationPools.All(p => p.Name != "VergunningenPool"))
                    {
                        Console.WriteLine("Creating VergunningenPool...");
                        ApplicationPool vergunningenPool = serverManager.ApplicationPools.Add("VergunningenPool");
                        vergunningenPool.ManagedRuntimeVersion = "v4.0";
                        vergunningenPool.AutoStart = true;
                        vergunningenPool.ManagedPipelineMode = ManagedPipelineMode.Integrated;
                        vergunningenPool.ProcessModel.IdentityType = ProcessModelIdentityType.SpecificUser;
                        vergunningenPool.ProcessModel.UserName = user.Name;
                        vergunningenPool.ProcessModel.Password = user.Password;
                    }

                    if (serverManager.ApplicationPools.All(p => p.Name != "DBAImportPool"))
                    {
                        Console.WriteLine("Creating DBAImportPool...");
                        ApplicationPool dbaImportPool = serverManager.ApplicationPools.Add("DBAImportPool");
                        dbaImportPool.ManagedRuntimeVersion = "v4.0";
                        dbaImportPool.AutoStart = true;
                        dbaImportPool.ManagedPipelineMode = ManagedPipelineMode.Integrated;
                        dbaImportPool.ProcessModel.IdentityType = ProcessModelIdentityType.SpecificUser;
                        dbaImportPool.ProcessModel.UserName = user.Name;
                        dbaImportPool.ProcessModel.Password = user.Password;
                    }

                    if (serverManager.ApplicationPools.All(p => p.Name != "PlanregisterPool"))
                    {
                        Console.WriteLine("Creating PlanregisterPool...");
                        ApplicationPool planregisterPool = serverManager.ApplicationPools.Add("PlanregisterPool");
                        planregisterPool.ManagedRuntimeVersion = "v4.0";
                    }

                    serverManager.CommitChanges();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error Adding Application Pools: " + ex.Message);
                }
            }
        }

        private static ImpersonatedUser GetImpersonatedUser()
        {
            ImpersonatedUser user = new ImpersonatedUser();
            user.GetName();
            user.GetPassword();
            return user;
        }

        private static ConfigurationElement FindElement(ConfigurationElementCollection collection, string elementTagName, params string[] keyValues)
        {
            foreach (ConfigurationElement element in collection)
            {
                if (String.Equals(element.ElementTagName, elementTagName, StringComparison.OrdinalIgnoreCase))
                {
                    bool matches = true;
                    for (int i = 0; i < keyValues.Length; i += 2)
                    {
                        object o = element.GetAttributeValue(keyValues[i]);
                        string value = null;
                        if (o != null)
                        {
                            value = o.ToString();
                        }
                        if (!String.Equals(value, keyValues[i + 1], StringComparison.OrdinalIgnoreCase))
                        {
                            matches = false;
                            break;
                        }
                    }
                    if (matches)
                    {
                        return element;
                    }
                }
            }
            return null;
        }

        private static bool UnlockSection(string sectionName)
        {
            try
            {
                System.Diagnostics.Process appCmdProc = new System.Diagnostics.Process
                {
                    StartInfo =
                    {
                        FileName = @"C:\Windows\System32\inetsrv\appcmd.exe",
                        Arguments = "unlock config /section:" + sectionName,
                        UseShellExecute = false,
                        RedirectStandardOutput = true
                    }
                };
                appCmdProc.Start();
                Console.WriteLine(appCmdProc.StandardOutput.ReadToEnd());
                //appCmdProc.WaitForExit();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error unlocking section " + sectionName + ": " + ex.Message);
                return false;
            }
        }
    }
}
