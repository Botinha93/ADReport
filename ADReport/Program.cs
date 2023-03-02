using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Security;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.IO;

namespace ADReport
{

    class Program
    {
 
        static void Main(string[] args)
        {
            /*Usuario.userWorkbookInit(Domain.GetComputerDomain().Name);
            GetADUsers();
            Usuario.Finish();*/
            config.init();
        }
        /*private static void GetADUsers()
        {
            
            using (var context = new PrincipalContext(ContextType.Domain, Domain.GetComputerDomain().Name))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    foreach (var result in searcher.FindAll())
                    {
                        DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                        new Usuario(de.Property("name"), de.Property("UserPrincipalName"), de.Property("lastLogon"), de.Property("Enabled"), de.Property("DisplayName"), de.Property("mail"), de.Property("CanonicalName"));
                    }
                }
            }
        }*/

        static public void dosearch(String[] properties,Principal p, operations name)
        {
            
            GenericWorkbook workbook = new GenericWorkbook(name, config.domain);
            using (var searcher = new PrincipalSearcher(p))
            {
                
                Boolean initiated = false;
                foreach (var result in searcher.FindAll())
                {
                    DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                    var dc = de.GetProperties(properties);
                    if (!initiated)
                    {
                        workbook.WorkbookInit(dc);
                        initiated = true;
                    }
                    workbook.WorkbookAdd(dc);
                }
            }
            workbook.Finish();
        }

        static public void dosearchD(String[] properties, DirectoryEntry root ,String filter, operations data)
        {

            GenericWorkbook workbook = new GenericWorkbook(data, config.domain);
            using (DirectorySearcher deSearch = new DirectorySearcher())
            {
                deSearch.SearchRoot = root;
                deSearch.Filter = filter;
                Boolean initiated = false;
                foreach (SearchResult de in deSearch.FindAll())
                {
                    var result = de.GetDirectoryEntry();
                    var dc = result.GetProperties(properties);
                    if (!initiated)
                    {
                        workbook.WorkbookInit(dc);
                        initiated = true;
                    }
                    workbook.WorkbookAdd(dc);
                }
            }
            workbook.Finish();
        }

    }
}
