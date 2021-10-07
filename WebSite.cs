using System;
using System.Collections.Generic;
using System.Text;

namespace IISManager
{
    public class WebSite
    {
        private string netVersion;
        private string appPoolName;
        private string siteId;
        private string siteName;
        private int port = 8080;
        public string ServerName { get; set; }
        public string Path { get; set; }
        public string AppPoolName { get => this.appPoolName ?? "DefaultAppPool"; set { appPoolName = value; } }
        public string MetaPath { get => "IIS://" + (this?.ServerName ?? "localhost") + "/W3SVC/"; }
        public string NetVersion { get => this.netVersion ?? ""; set { netVersion = value; } }
        public string SiteId { get => this?.siteId ?? "1"; set { siteId = value; } }
        public string DirName { get; set; }
        public string SiteName { get => this.siteName ?? "Default Web Site"; set { siteName = value; } }
        public int Port { get => this.port; set => this.port = value; }
        //Actions
        public bool EditPool { get; set; }
        public bool EditWebSite { get; set; }
        public bool AssignSiteToPool { get; set; }
        public bool Https { get; set; }
        public string CertificatePath { get; set; }
        public string CertificateName { get; set; }
    }

}
