using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace LoadMail
{
    public partial class ThisAddIn
    {
        #region Instance Variables
        Outlook.Application m_Application;
        Outlook.Explorers m_Explorers;
        Outlook.Inspectors m_Inspectors;
        public stdole.IPictureDisp m_pictdisp = null;
        // Ribbon UI reference.
        internal static Office.IRibbonUI m_Ribbon;
        #endregion


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

        }

        #endregion

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            m_Application = new Outlook.Application();
            m_Inspectors = this.Application.Inspectors;

            m_Inspectors.NewInspector +=
                new Outlook.InspectorsEvents_NewInspectorEventHandler(GetMailItemEntryId_Click);


            this.Application.ItemContextMenuDisplay +=
                new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);

            this.Application.NewMail +=
                new Outlook.ApplicationEvents_11_NewMailEventHandler(EmailArrived);

            this.Application.ItemLoad +=
                new Outlook.ApplicationEvents_11_ItemLoadEventHandler(GetMailItemLocation);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 激活邮件时
        /// </summary>
        /// <param name="Item"></param>
        private void GetMailItemLocation(object Item)
        {
            if (m_Application.ActiveExplorer().CurrentFolder != null)
            {
                var mailItem = m_Application.ActiveExplorer().CurrentFolder;
                Outlook.MAPIFolder oFolder = (Outlook.MAPIFolder)mailItem;
                //MessageBox.Show("打开位于 {" + oFolder.FolderPath.ToString() + "} 的邮件列表");
            }

        }

        /// <summary>
        /// 有新邮件时
        /// </summary>
        private void EmailArrived()
        {
            //MessageBox.Show("hi,你有新邮件！");
        }

        /// <summary>
        /// 打开新窗口（阅读 或者 新建 邮件）时
        /// </summary>
        /// <param name="Inspector"></param>
        private void GetMailItemEntryId_Click(Outlook.Inspector Inspector)
        {
            //MessageBox.Show("打开新窗口");
            //throw new NotImplementedException();
            Outlook.MailItem objMailItem = Inspector.CurrentItem;

        }

        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {

            if (Selection.Count == 1)
            {
                if (Selection[1] is Outlook.MailItem)
                {
                    this.CustomContextMenu(CommandBar, Selection);
                }
            }

            if (Selection.Count > 1)
            {
                this.CustomContextBatchMenu(CommandBar, Selection);
            }
        }



        Outlook.Selection CurrentSelection;
        private void CustomContextMenu(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            Office.CommandBarButton customContextMenuButton =
                          (Office.CommandBarButton)CommandBar.Controls.Add
                (Office.MsoControlType.msoControlButton
                                     , Type.Missing
                                     , Type.Missing
                                     , Type.Missing
                                     , true);
            CurrentSelection = Selection;

            customContextMenuButton.Caption = "导入PPManager";
            customContextMenuButton.FaceId = 446;
            customContextMenuButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(customContextMenuTag_Click);



        }

        private void CustomContextBatchMenu(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            Office.CommandBarButton customContextBatchMenuButton =
              (Office.CommandBarButton)CommandBar.Controls.Add
    (Office.MsoControlType.msoControlButton
                         , Type.Missing
                         , Type.Missing
                         , Type.Missing
                         , true);
            CurrentSelection = Selection;

            customContextBatchMenuButton.Caption = "批量导入PPManager";
            customContextBatchMenuButton.FaceId = 447;
            customContextBatchMenuButton.Click += customContextBatchMenuButton_Click;

        }

        void customContextBatchMenuButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            string error = string.Empty;
            StringBuilder sbr = new StringBuilder();
            if (CurrentSelection.Count > 0)
            {
                if (MessageBox.Show("本次要处理{" + CurrentSelection.Count + "}封邮件，可能需要等待较长时间。\n\r 在服务器返回结果之前请勿进行其他操作。\n\r 确定继续吗？", "友情提醒", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    sbr.AppendFormat("本次一共处理了{0}封邮件", CurrentSelection.Count);
                    sbr.AppendLine();
                    for (int i = 0; i < CurrentSelection.Count; i++)
                    {
                        if (CurrentSelection[i + 1] is Outlook.MailItem)
                        {
                            var item = (Outlook.MailItem)CurrentSelection[i + 1];

                            if (Process.BatchToAPI(item, out error))
                            {
                                sbr.AppendFormat("[成功]{0}", item.Subject);
                                sbr.AppendLine();
                            }
                            else
                            {
                                sbr.AppendFormat("[失败]{0}{1}", item.Subject, error);
                                sbr.AppendLine();
                            }
                        }
                    }

                    MessageBox.Show("导入结束，结果：\r\n" + sbr.ToString());
                }

            }
            else
            {
                MessageBox.Show("请选择邮件要处理的邮件，可选多个");
            }
        }


        void customContextMenuTag_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (CurrentSelection.Count == 1)
            {
                var item = (Outlook.MailItem)CurrentSelection[1];

                Process.ToAPI(item);
            }
            else
            {
                MessageBox.Show("选择一封邮件。");
            }
        }




    }
}
