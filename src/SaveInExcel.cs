using System;
using Microsoft.Office.Interop.Excel;

namespace EmailFromPDF
{
    public class SaveInExcel
    {
        public SaveInExcel(string[] text, string path)
        {
            var oXl = new Application();

            _Workbook oWb = oXl.Workbooks.Add("");
            var oSheet = (_Worksheet)oWb.ActiveSheet;

            for (var r = 0; r < text.Length; r++)
            {
                oSheet.Cells[r + 1, 1] = text[r];
            }

            oXl.Visible = false;
            oXl.UserControl = false;
            var date = DateTime.Now.ToString("yy-MM-dd HH_mm_ss");
            oWb.SaveAs(path + "\\Emails_" + text.Length + "_" + date + ".xls", XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWb.Close();
        }
    }
}
