using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADReport
{
    class operations
    {
        public String[] props;
        public String name;
        public String filepath;
        public String op;
        public String filter;
    }
    class config
    {
        static public string domain = Domain.GetComputerDomain().Name;
        static public void init()
        {
            string jsonFilePath = @"config.json";
            string json = File.ReadAllText(jsonFilePath);
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            List<operations> usr = serializer.Deserialize<List<operations>>(json);
            var context = new PrincipalContext(ContextType.Domain, config.domain);
            String[] d = domain.Split('.');
            string localizedDomain = "LDAP://";
            for (var i = 0; i < d.Length; i++)
            {
                localizedDomain += "DC=" + d[i];
                if(i < (d.Length - 1))
                {
                    localizedDomain += ",";
                }
            }
            DirectoryEntry contextd = new DirectoryEntry(localizedDomain);
            foreach (var opreation in usr)
            {
                switch (opreation.op){
                    case "users":
                        Program.dosearch(opreation.props, new UserPrincipal(context), opreation);
                        break;
                    case "Groups":
                        Program.dosearch(opreation.props, new GroupPrincipal(context), opreation);
                        break;
                    case "ou":
                        Program.dosearchD(opreation.props, contextd, "(objectCategory=organizationalUnit)", opreation);
                        break;
                    case "Computer":
                        Program.dosearch(opreation.props, new ComputerPrincipal(context), opreation);
                        break;
                    case "Trust":
                        Program.dosearchD(opreation.props, contextd, "(objectClass=trustedDomain)", opreation);
                        break;
                    case "Subnets":
                        Program.dosearchD(opreation.props, contextd, "(objectClass=subnet)", opreation);
                        break;
                    case "DomainController":
                        Program.dosearchD(opreation.props, contextd, "(primaryGroupID=516)", opreation);
                        break;
                    case "Custom":
                        Program.dosearchD(opreation.props, contextd, opreation.filter, opreation);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
