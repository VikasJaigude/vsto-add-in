using Microsoft.Office.Tools.Ribbon;

namespace OutlookAddIn2
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 dlg = new Form1();
            dlg.ShowDialog();
        }
    }
}
