using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace fncGTD4Outlook
{
    public partial class RibbonCompose
    {
        private void RibbonCompose_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnDelegarEnviar_Click(object sender, RibbonControlEventArgs e)
        {
            frmDelegarEnviar f = new frmDelegarEnviar();
            f.Show();
            f = null;
        }
    }
}
