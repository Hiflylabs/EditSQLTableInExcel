using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {            
            Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
