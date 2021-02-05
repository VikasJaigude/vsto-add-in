using Microsoft.Office.Tools.Ribbon;
using Microsoft.Web.WebView2.Core;
using System;

namespace OutlookAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Form1 dlg = new Form1();
            //dlg.ShowDialog();
            if (IsWebViewVersionInstalled())
            {
                Form1 dlg = new Form1();
                dlg.ShowDialog();
            }
            else
            {
                Form2 dlg = new Form2();
                dlg.ShowDialog();
            }

        }

        private bool IsWebViewVersionInstalled()
        {
            string versionNo = null;
            Version asmVersion = null;

            try
            {
                versionNo = CoreWebView2Environment.GetAvailableBrowserVersionString();

                asmVersion = typeof(CoreWebView2Environment).Assembly.GetName().Version;

                //if (ver.Build >= asmVersion.Build)
                if (asmVersion.Build > 0)
                    return true;
            }
            catch { }

            return false;
        }
    }
}
