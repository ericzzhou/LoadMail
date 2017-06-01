using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace LoadMail
{
    public partial class Rbb_PPManager
    {

        private void Rbb_PPManager_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnLoadMail_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Outlook.Application objApplication = Globals.ThisAddIn.Application;
                Outlook.Inspector objInspector = objApplication.ActiveInspector();


                Outlook.MailItem objMailItem = objInspector.CurrentItem;
                if (objMailItem != null)
                {
                    Process.ToAPI(objMailItem);
                    //string mailBody = objMailItem.Body;
                    //MessageBox.Show("附件个数：" + objMailItem.Attachments.Count);
                    //for (int i = 0; i < objMailItem.Attachments.Count; i++)
                    //{
                    //    string savePaht =@"E:\Work\Project\ppmanager\01.SourceCode\plug-in";
                    //    objMailItem.Attachments[i + 1].SaveAsFile(savePaht);
                    //    //MssageBox.Show();
                    //}

                    //MessageBox.Show(mailBody);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


    }
}
