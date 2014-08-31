using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDna.ExcelDnaTools
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            // string tabLabel = ExcelDnaUtil.ExcelVersion >= 15.0 ? "EXCEL-DNA" : "Excel-DNA";
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
    <ribbon>
        <tabs>
            <tab id='ExcelDna' getLabel='GetLabel' keytip='-'>
            <group id='IdeManager' label='IDE Manager' >
                <button id='ShowIde' label='Show IDE' size='large' onAction='OnShowIde' keytip='S' />
                <button id='HideIde' label='Hide IDE' size='large' onAction='OnHideIde' keytip='H' />
                <button id='ShutdownIde' label='Shutdown IDE' size='large' onAction='OnShutdownIde' keytip='U' />
            </group >
            </tab>
        </tabs>
    </ribbon>
</customUI>";
        }

        public string GetLabel(IRibbonControl control)
        {
            return ExcelDnaUtil.ExcelVersion >= 15.0 ? "EXCEL-DNA" : "Excel-DNA";
        }

        public void OnShowIde(IRibbonControl control)
        {
            VsConnection.Show();
        }

        public void OnHideIde(IRibbonControl control)
        {
            VsConnection.Hide();
        }

        public void OnShutdownIde(IRibbonControl control)
        {
            VsConnection.Shutdown();
        }
    }
}
