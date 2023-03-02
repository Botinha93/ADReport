using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADReport
{
    class Usuario
    {
        static IList RegUsuario;
        String name;
        String UserPrincipalName;
        String lastlogondate;
        bool Enabled;
        String DisplayName;
        String mail;
        String CanonicalName;
        static string domain;
        static ClosedXML.Excel.XLWorkbook workbook;
        static ClosedXML.Excel.IXLWorksheet worksheet;
        public Usuario(String name,String UserPrincipalName, String lastlogondate,String Enabled,String DisplayName,String mail,String CanonicalName)
        {
            this.lastlogondate = lastlogondate;
            this.name = name;
            this.UserPrincipalName = UserPrincipalName;
            this.Enabled = (Enabled.Contains("false") ? false : true);
            this.mail = mail;
            this.DisplayName = DisplayName;
            this.CanonicalName = CanonicalName;
            RegUsuario.Add(this);
            makeLine(RegUsuario.Count + 1);
        }
        private void makeLine(int pos)
        {
            worksheet.Cell("A" + pos).Value = name;
            worksheet.Cell("B" + pos).Value = UserPrincipalName;
            worksheet.Cell("C" + pos).Value = lastlogondate;
            worksheet.Cell("D" + pos).Value = Enabled;
            worksheet.Cell("E" + pos).Value = DisplayName;
            worksheet.Cell("F" + pos).Value = mail;
            worksheet.Cell("G" + pos).Value = CanonicalName;
            worksheet.Cell("H" + pos).Value = Usuario.domain;
        }
        static public void userWorkbookInit(string domain)
        {
            workbook = new ClosedXML.Excel.XLWorkbook();
            RegUsuario = new List<Usuario>();
            Usuario.domain = domain;
            worksheet = workbook.Worksheets.Add("Usuarios");
            worksheet.Cell("A1").Value = "name";
            worksheet.Cell("B1").Value = "UserPrincipalName";
            worksheet.Cell("C1").Value = "lastlogondate";
            worksheet.Cell("D1").Value = "Enabled";
            worksheet.Cell("E1").Value = "DisplayName";
            worksheet.Cell("F1").Value = "mail";
            worksheet.Cell("G1").Value = "CanonicalName";
            worksheet.Cell("H1").Value = "Domain";
        }
        public static void Finish()
        {
            workbook.SaveAs("AllUserAD - Usuários.xlsx");
        }
    }
}
