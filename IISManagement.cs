using Microsoft.Web.Administration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Principal;
using System.Text;

namespace IISManager
{
    public static class IISManagement
    {
        private static ServerManager serverManager = new ServerManager();

        public static List<string> GetSitesNames()
        {
            SiteCollection sites = serverManager.Sites;
            var sitesNames = new List<string>();
            foreach (var site in sites)
            {
                sitesNames.Add(site.Name);
            }
            return sitesNames;
        }

        public static List<List<string>> GetProcessesIds()
        {
            var pools = serverManager.ApplicationPools;
            var info = new List<List<string>>();
            foreach(var pool in pools)
            {
                info.Add(pool.WorkerProcesses.Select(x => $"{x.AppPoolName} {x.ProcessId}").ToList());
            }
            return info;
        }

        public static List<ApplicationPool> GetPools()
        {
            var pools = serverManager.ApplicationPools;
            var poolsInfo = new List<ApplicationPool>();
            foreach (var pool in pools)
            {
                poolsInfo.Add(pool);
            }
            return poolsInfo;
        }

        public static string DeletePoolTempFiles(WebSite poolInfo)
        {
            var poolName = poolInfo.AppPoolName;
            var poolNames = serverManager.ApplicationPools.Select(x => x.Name).ToList();
            var tempPath = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            var poolFolder = Path.Combine(tempPath, poolName);
            var directory = new DirectoryInfo(poolFolder);
            if(directory.Exists)
            {
                directory.Delete(true);
            }
            
            return "OK";
        }
        public static string GetPoolTempPath(string userName)
        {
            var userSID = GetUserSID(userName);
            var keyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + userSID;

            var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(keyPath);
            if (key == null)
            {
                //handle error
                return null;
            }

            var profilePath = key.GetValue("ProfileImagePath") as string;
            return profilePath;
        }
        public static string GetUserSID(string userName)
        {
            try
            {
                NTAccount f = new NTAccount(userName);
                SecurityIdentifier s = (SecurityIdentifier)f.Translate(typeof(SecurityIdentifier));
                return s.ToString();
            }
            catch
            {
                return null;
            }
        }

        public static string Test()
        {
            var poolName = "TestHook";
            var path = GetPoolTempPath(poolName);
            var tempPath = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            var poolFolder = Path.Combine(tempPath, poolName);
            var directory = new DirectoryInfo(poolFolder);
            return path;
        }
        public static string ChangeIISSettings(WebSite siteInfo)
        {
            var answer = new StringBuilder("");

            if (siteInfo.EditPool)
            {
                answer.Append(CreatePool(siteInfo));
            }
            if (siteInfo.EditWebSite)
            {
                answer.Append(CreateWebSite(siteInfo));
            }
            if (siteInfo.AssignSiteToPool)
            {
                answer.Append(AssignDirToPool(siteInfo));
            }
            answer.Append("Finished.");
            return answer.ToString();
        }


        public static string DeleteIISObjects(WebSite siteInfo)
        {
            var answer = new StringBuilder("");

            if (siteInfo.EditPool == true)
            {
                answer.Append(DeletePool(siteInfo));
            }
            if (siteInfo.EditWebSite == true)
            {
                answer.Append(DeleteWebSite(siteInfo));
            }
            answer.Append("Finished.");
            return answer.ToString();
        }
        public static string DeletePool(WebSite poolInfo)
        {
            var appPools = serverManager.ApplicationPools;
            DeletePoolTempFiles(poolInfo);
            try
            {
                appPools.FirstOrDefault(x => x.Name == poolInfo.AppPoolName).Delete();
                serverManager.CommitChanges();
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return $"Pool {poolInfo.AppPoolName} has been Deleted";
        }
        public static string DeleteWebSite(WebSite siteInfo)
        {
            var sites = serverManager.Sites;
            try
            {
                sites.FirstOrDefault(x => x.Name == siteInfo.SiteName).Delete();
                serverManager.CommitChanges();
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return $"WebSite {siteInfo.SiteName} has been Deleted";
        }

        public static string CreatePool(WebSite poolInfo)
        {
            var appPools = serverManager.ApplicationPools;
            if (appPools.Any(x => x.Name == poolInfo.AppPoolName))
            {
                return JsonConvert.SerializeObject(new ArgumentException("Pool Already Exists"));
                //return $"Pool {poolInfo.AppPoolName} already have been created";
            }

            try
            {
                serverManager.ApplicationPools.Add(poolInfo.AppPoolName);
                var pool = serverManager.ApplicationPools[poolInfo.AppPoolName];
                pool.ManagedRuntimeVersion = poolInfo.NetVersion;
                serverManager.CommitChanges();
                return "Creating Pool - OK";
            }
            catch (Exception e)
            {
                return JsonConvert.SerializeObject(e);
            }
        }

        public static string CreateWebSite(WebSite siteInfo)
        {
            var sites = serverManager.Sites;
            if (sites.Any(x => x.Name == siteInfo.SiteName))
            {
                if (siteInfo.Https)
                {
                    var site = sites.FirstOrDefault(x => x.Name == siteInfo.SiteName);
                    var certInfo = GenerateCertificate(siteInfo);

                    var binding = site.Bindings.Add($"*:{siteInfo.Port}:", certInfo, "My");
                    binding.Protocol = "https";
                    serverManager.CommitChanges();
                }
                return $"WebSite {siteInfo.AppPoolName} already have been created";
            }
            try
            {
                if (!System.IO.Directory.Exists(siteInfo.Path))
                {
                    System.IO.Directory.CreateDirectory(siteInfo.Path);
                }
                Site site;
                if (siteInfo.Https)
                {

                    var certInfo = GenerateCertificate(siteInfo);
                    site = sites.Add(siteInfo.SiteName, $"*:{siteInfo.Port}:", siteInfo.Path, certInfo);
                }
                else
                {
                    site = sites.Add(siteInfo.SiteName, siteInfo.Path, siteInfo.Port);
                }

                serverManager.CommitChanges();
                return $"Site {siteInfo.SiteName} have been created";
            }
            catch (Exception e)
            {
                return JsonConvert.SerializeObject(e);
            }
        }

        public static byte[] GenerateCertificate(WebSite info)
        {
            X509Store store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            store.Open(OpenFlags.OpenExistingOnly | OpenFlags.ReadWrite);

            var rsa = RSA.Create(4096);
            var req = new CertificateRequest($"cn={info.ServerName}", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            var cert = req.CreateSelfSigned(DateTimeOffset.Now, DateTimeOffset.Now.AddYears(1));
            if (!System.IO.Directory.Exists(info.CertificatePath))
            {
                System.IO.Directory.CreateDirectory(info.CertificatePath);
            }
            var path = Path.Combine(info.CertificatePath, info.CertificateName);
            var certInfo = cert.Export(X509ContentType.Pfx, "password");
            System.IO.File.WriteAllBytes($"{path}.pfx", certInfo);

            System.IO.File.WriteAllText($"{path}.cer",
                "-----BEGIN CERTIFICATE-----\r\n"
                + Convert.ToBase64String(cert.Export(X509ContentType.Cert), Base64FormattingOptions.InsertLineBreaks)
                + "\r\n-----END CERTIFICATE-----");
            var certificate = new X509Certificate2($"{path}.pfx", "password", X509KeyStorageFlags.UserKeySet);
            store.Add(certificate);
            store.Close();

            return certificate.GetCertHash();
        }

        public static string AssignDirToPool(WebSite siteInfo)
        {
            if (serverManager.Sites.Any(x => x.Name == siteInfo.SiteName))
            {
                var site = serverManager.Sites.FirstOrDefault(x => x.Name == siteInfo.SiteName);
                if (site.ApplicationDefaults.ApplicationPoolName == siteInfo.AppPoolName)
                {
                    return $"WebSite {siteInfo.SiteName} aready assigned to Pool {siteInfo.AppPoolName}";
                }
                else
                {
                    site.ApplicationDefaults.ApplicationPoolName = siteInfo.AppPoolName;
                    serverManager.CommitChanges();
                }
            }
            else
            {
                return $"WebSite {siteInfo.SiteName} doesn't exist. Create First!";
            }


            return $"Pool {siteInfo.AppPoolName} have been assigned to website {siteInfo.SiteName}";
        }
    }
}
