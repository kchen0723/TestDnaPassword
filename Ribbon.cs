using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace TestDnaPassword
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon, IExcelAddIn
    {
        public override string GetCustomUI(string RibbonID)
        {
            string result = @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
  <ribbon>
    <tabs>
      <tab id=""MI"" label=""TestPassword"" >
        <group id=""ContentGroup"">
          <button id=""aboutButton"" label=""About"" size=""large""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
            return result;
        }
        public void AutoClose()
        {
        }

        public void AutoOpen()
        {
            ((Application)ExcelDnaUtil.Application).WorkbookBeforeClose += Ribbon_WorkbookBeforeClose;
        }

        private void Ribbon_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            Console.WriteLine("workbook Closing now");
        }
    }
}
