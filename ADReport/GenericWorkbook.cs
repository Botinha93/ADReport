using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADReport
{
    class GenericWorkbook
    {
        ClosedXML.Excel.XLWorkbook workbook;
        operations workbookName;
        ClosedXML.Excel.IXLWorksheet worksheet;
        int currentRow = 1;
        string domain;
        public GenericWorkbook(operations workbookName, string domain)
        {
            workbook = new ClosedXML.Excel.XLWorkbook();
            this.workbookName = workbookName;
            this.domain = domain;
        }
        public void WorkbookInit(Dictionary<string, string> dic)
        {
            workbook = new ClosedXML.Excel.XLWorkbook();
            worksheet = workbook.Worksheets.Add(workbookName.name);
            int cell = 1;
            foreach (var d in dic)
            {
                worksheet.Row(currentRow).Cell(cell).Value = d.Key;
                cell++;
            }
            worksheet.Row(currentRow).Cell(cell).Value = "domain";
            currentRow++;
        }
        public void WorkbookAdd(Dictionary<string, string> dic)
        {
            int cell = 1;
            foreach (var d in dic)
            {
                worksheet.Row(currentRow).Cell(cell).Value = d.Value;
                cell++;
            }
            worksheet.Row(currentRow).Cell(cell).Value = domain;
            currentRow++;
        }
        public void Finish()
        {
            try { workbook.SaveAs(workbookName.filepath + workbookName.name + ".xlsx"); } catch { }

        }
    }
}
