using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

/// <summary>
/// VisualSVN Server管理的Excel插件
/// 
/// 伍鲜
/// </summary>
namespace SVNManagementAddIn
{
    /// <summary>
    /// 插件实现代码
    /// </summary>
    public partial class ExcelRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Environment.GetEnvironmentVariable("VISUALSVN_SERVER");

            btnDeleteRepositories.Enabled = false;
            btnDeleteGroups.Enabled = false;
        }

        #region 调用系统接口
        private enum ShowCommands : int
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }

        [DllImport("shell32.dll")]
        static extern IntPtr ShellExecute(
            IntPtr hwnd,
            string lpOperation,
            string lpFile,
            string lpParameters,
            string lpDirectory,
            ShowCommands showCmd);

        #endregion

        #region 邮件发送

        private enum EmailMode
        {
            Mailto,
            MailMessage,
            Outlook
        }

        private EmailMode mode = EmailMode.MailMessage;

        #endregion

        #region 仓库管理

        /// <summary>
        /// 创建仓库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateRepositories_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["创建仓库"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            while (++row <= rows)
            {
                if (sheet.Cells[row, 6].Value == null || "".Equals(sheet.Cells[row, 6].Value.Trim()))
                {
                    bool result = SVNHelper.CreateRepository(sheet.Cells[row, 1].Value);
                    if (result)
                    {
                        if (sheet.Cells[row, 2].Value == "独立仓库")
                        {
                            SVNHelper.CreateRepositoryFolders(Convert.ToString(sheet.Cells[row, 1].Value), "trunk,branches,tags".Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries));
                        }

                        if (client != null)
                        {
                            if (mode == EmailMode.MailMessage)
                            {
                                MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 5].Value), Convert.ToString(sheet.Cells[row, 4].Value));
                                MailMessage message = new MailMessage(from, to);

                                message.Subject = subject;

                                message.Bcc.Add(adminEmail);

                                message.Attachments.Add(imageInEmailHead);
                                message.Attachments.Add(imageInEmailImgBlue);
                                message.Attachments.Add(imageInEmailFooter);

                                message.IsBodyHtml = true;

                                message.Body = File.ReadAllText(
                                    AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/RepositoryCreate.html")
                                    .Replace("${Username}", Convert.ToString(sheet.Cells[row, 4].Value))
                                    .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                    .Replace("${SVNRepositoryName}", Convert.ToString(sheet.Cells[row, 1].Value))
                                    .Replace("${SVNRepositoryDesc}", Convert.ToString(sheet.Cells[row, 3].Value))
                                    .Replace("${AdminEmail}", adminEmail)
                                    .Replace("${AdminName}", adminName)
                                    .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                    );

                                client.Send(message);
                            }
                        }
                        sheet.Cells[row, 6] = "是";
                        sheet.Cells[row, 7] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 创建仓库目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateRepositoriesFolders_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["创建仓库目录"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 3].Value == null || "".Equals(sheet.Cells[row, 3].Value.Trim()))
                {
                    bool result = SVNHelper.CreateRepositoryFolders(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value).Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries));
                    if (result)
                    {
                        sheet.Cells[row, 3] = "是";
                        sheet.Cells[row, 4] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 重命名仓库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRenameRepositories_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["重命名仓库"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 3].Value == null || "".Equals(sheet.Cells[row, 3].Value.Trim()))
                {
                    bool result = SVNHelper.RenameRepository(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value));
                    if (result)
                    {
                        sheet.Cells[row, 3] = "是";
                        sheet.Cells[row, 4] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 删除仓库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteRepositories_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["删除仓库"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 2].Value == null || "".Equals(sheet.Cells[row, 2].Value.Trim()))
                {
                    bool result = SVNHelper.DeleteRepository(Convert.ToString(sheet.Cells[row, 1].Value));
                    if (result)
                    {
                        sheet.Cells[row, 2] = "是";
                        sheet.Cells[row, 3] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        #endregion

        #region 组管理

        /// <summary>
        /// 创建组
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateGroups_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["创建组"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 3].Value == null || "".Equals(sheet.Cells[row, 3].Value.Trim()))
                {
                    bool result = SVNHelper.CreatGroup(Convert.ToString(sheet.Cells[row, 1].Value));
                    if (result)
                    {
                        sheet.Cells[row, 3] = "是";
                        sheet.Cells[row, 4] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 设置组员
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetMembers_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["设置组员"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value.Trim()))
                {
                    bool result = SVNHelper.SetGroupMembers(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value), Convert.ToString(sheet.Cells[row, 3].Value));
                    if (result)
                    {
                        sheet.Cells[row, 4] = "是";
                        sheet.Cells[row, 5] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 删除组
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteGroups_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["删除组"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 2].Value == null || "".Equals(sheet.Cells[row, 2].Value.Trim()))
                {
                    bool result = SVNHelper.DeleteGroup(Convert.ToString(sheet.Cells[row, 1].Value));
                    if (result)
                    {
                        sheet.Cells[row, 2] = "是";
                        sheet.Cells[row, 3] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        #endregion

        #region 用户管理

        /// <summary>
        /// 根据用户名称生成用户密码
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        private string GeneratePassword(string username)
        {
            // return DESEncrypt.DesEncrypt((username + "ABCDEFGHIJKLMNOPQRSTUVWXYZ").Substring(0, 16));

            Random rd = new Random();
            Regex r = new Regex(@"[a-zA-Z]+");
            MatchCollection mc = r.Matches(username);
            StringBuilder builder = new StringBuilder();

            for (int i = 0; i < mc.Count; i++)
            {
                builder.Append(mc[i].Value);
            }

            string passString = builder.ToString() + Guid.NewGuid().ToString().Replace("-", "") + builder.ToString();

            builder = new StringBuilder();

            for (int i = 0; i < 32; i++)
            {
                builder.Append(passString.Substring(rd.Next(0, passString.Length), 1));
            }

            return builder.ToString().ToUpper();
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateUsers_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string managerName = sheet.Cells[1, 3].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["创建用户"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            StringBuilder builder = new StringBuilder();

            while (++row <= rows)
            {
                if (sheet.Cells[row, 6].Value == null || "".Equals(sheet.Cells[row, 6].Value.Trim()))
                {
                    if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value))
                    {
                        sheet.Cells[row, 4].Value = Convert.ToString(sheet.Cells[row, 3].Value).Split('@')[0].ToLower().Replace("_", ".");
                    }
                    sheet.Cells[row, 5].Value = GeneratePassword(Convert.ToString(sheet.Cells[row, 4].Value));
                    bool result = SVNHelper.CreateUser(Convert.ToString(sheet.Cells[row, 4].Value), Convert.ToString(sheet.Cells[row, 5].Value));
                    if (result)
                    {
                        if (client != null)
                        {
                            builder.Append("名称：").Append(sheet.Cells[row, 1].Value).Append("，&nbsp;");
                            builder.Append("邮箱：").Append(sheet.Cells[row, 3].Value).Append("，&nbsp;");
                            builder.Append("账号：").Append(sheet.Cells[row, 4].Value).Append("，&nbsp;");
                            builder.Append("密码：").Append(sheet.Cells[row, 5].Value).Append("<br />");

                            if (mode == EmailMode.MailMessage)
                            {
                                MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 1].Value));
                                MailMessage message = new MailMessage(from, to);

                                message.Subject = subject;

                                message.Attachments.Add(imageInEmailHead);
                                message.Attachments.Add(imageInEmailImgBlue);
                                message.Attachments.Add(imageInEmailFooter);

                                message.IsBodyHtml = true;

                                message.Body = File.ReadAllText(
                                    AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserCreate.html")
                                    .Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value))
                                    .Replace("${SVNUsername}", Convert.ToString(sheet.Cells[row, 4].Value))
                                    .Replace("${SVNPassword}", Convert.ToString(sheet.Cells[row, 5].Value))
                                    .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                    .Replace("${AdminEmail}", adminEmail)
                                    .Replace("${AdminName}", adminName)
                                    .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                    );

                                client.Send(message);
                            }
                            else if (mode == EmailMode.Outlook)
                            { // 使用Outlook已经登录的账号进行发送
                                Outlook.Application olApp = new Outlook.Application();
                                Outlook.MailItem mailItem = (Outlook.MailItem)olApp.CreateItem(Outlook.OlItemType.olMailItem);
                                mailItem.To = sheet.Cells[row, 3].Value;
                                mailItem.Subject = subject;

                                mailItem.Attachments.Add(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
                                mailItem.Attachments.Add(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
                                mailItem.Attachments.Add(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");

                                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                                mailItem.HTMLBody = File.ReadAllText(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserCreate.html").Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value)).Replace("${SVNUsername}", Convert.ToString(sheet.Cells[row, 4].Value)).Replace("${SVNPassword}", Convert.ToString(sheet.Cells[row, 5].Value)).Replace("${AdminEmail}", adminEmail).Replace("${AdminName}", adminName).Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd"));

                                mailItem.Send();
                                mailItem = null;
                                olApp = null;
                            }
                        }
                        sheet.Cells[row, 6] = "是";
                        sheet.Cells[row, 7] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }

            if (builder.Length > 0 && client != null && manager != null && !"".Equals(manager))
            {
                if (mode == EmailMode.MailMessage)
                {
                    MailAddress to = new MailAddress(manager);
                    MailMessage message = new MailMessage(from, to);

                    message.Bcc.Add(adminEmail);

                    message.Subject = subject;

                    message.Attachments.Add(imageInEmailHead);
                    message.Attachments.Add(imageInEmailImgBlue);
                    message.Attachments.Add(imageInEmailFooter);

                    message.IsBodyHtml = true;

                    message.Body = File.ReadAllText(
                        AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserCreateManager.html")
                        .Replace("${ManagerName}", managerName)
                        .Replace("${SVNUserlist}", builder.ToString())
                        .Replace("${SVNRepositoryRoot}", repositoryRoot)
                        .Replace("${AdminEmail}", adminEmail)
                        .Replace("${AdminName}", adminName)
                        .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                        );

                    client.Send(message);
                }
            }
        }

        /// <summary>
        /// 设置密码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetPassword_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["设置用户密码"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            while (++row <= rows)
            {
                if (sheet.Cells[row, 6].Value == null || "".Equals(sheet.Cells[row, 6].Value.Trim()))
                {
                    if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value))
                    {
                        sheet.Cells[row, 4].Value = Convert.ToString(sheet.Cells[row, 3].Value).Split('@')[0].ToLower().Replace("_", ".");
                    }
                    sheet.Cells[row, 5].Value = GeneratePassword(Convert.ToString(sheet.Cells[row, 4].Value));

                    bool result = SVNHelper.SetPassword(Convert.ToString(sheet.Cells[row, 4].Value), Convert.ToString(sheet.Cells[row, 5].Value));
                    if (result)
                    {
                        if (client != null)
                        {
                            if (mode == EmailMode.MailMessage)
                            {
                                MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 1].Value));
                                MailMessage message = new MailMessage(from, to);

                                message.Subject = subject;

                                message.Attachments.Add(imageInEmailHead);
                                message.Attachments.Add(imageInEmailImgBlue);
                                message.Attachments.Add(imageInEmailFooter);

                                message.IsBodyHtml = true;

                                message.Body = File.ReadAllText(
                                    AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserPassword.html")
                                    .Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value))
                                    .Replace("${SVNUsername}", Convert.ToString(sheet.Cells[row, 4].Value))
                                    .Replace("${SVNPassword}", Convert.ToString(sheet.Cells[row, 5].Value))
                                    .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                    .Replace("${AdminEmail}", adminEmail)
                                    .Replace("${AdminName}", adminName)
                                    .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                    );

                                client.Send(message);
                            }
                        }
                        sheet.Cells[row, 6] = "是";
                        sheet.Cells[row, 7] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 设置用户组
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetGroups_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["设置用户组"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value.Trim()))
                {
                    SVNHelper.SetMemberGroup(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value), Convert.ToString(sheet.Cells[row, 3].Value));

                    sheet.Cells[row, 4] = "是";
                    sheet.Cells[row, 5] = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }
        }

        /// <summary>
        /// 用户邮箱验证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEmailVerify_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["验证用户邮箱"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 1;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            while (++row <= rows)
            {
                if (sheet.Cells[row, 5].Value == null || "".Equals(sheet.Cells[row, 5].Value))
                {
                    if (client != null)
                    {
                        if (mode == EmailMode.MailMessage)
                        {
                            MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 1].Value));
                            MailMessage message = new MailMessage(from, to);

                            message.Subject = subject;

                            message.Attachments.Add(imageInEmailHead);
                            message.Attachments.Add(imageInEmailImgBlue);
                            message.Attachments.Add(imageInEmailFooter);

                            message.IsBodyHtml = true;

                            message.Body = File.ReadAllText(
                                AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserEmailVerify.html")
                                .Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value))
                                .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                .Replace("${AdminEmail}", adminEmail)
                                .Replace("${AdminName}", adminName)
                                .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                );
                            client.Send(message);
                        }
                    }
                    sheet.Cells[row, 5] = "是";
                    sheet.Cells[row, 6] = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }
        }

        /// <summary>
        /// 定期刷新用户密码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnResetPassword_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["重置用户密码"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 1;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            while (++row <= rows)
            {
                if (sheet.Cells[row, 7].Value == null || "".Equals(sheet.Cells[row, 7].Value))
                {
                    if (sheet.Cells[row, 6].Value == "是")
                    { // 需要定期修改密码的
                        if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value))
                        {
                            sheet.Cells[row, 4].Value = Convert.ToString(sheet.Cells[row, 3].Value).Split('@')[0].ToLower().Replace("_", ".");
                        }
                        sheet.Cells[row, 5].Value = GeneratePassword(Convert.ToString(sheet.Cells[row, 4].Value));

                        bool result = SVNHelper.SetPassword(Convert.ToString(sheet.Cells[row, 4].Value), Convert.ToString(sheet.Cells[row, 5].Value));
                        if (result)
                        {
                            if (client != null)
                            {
                                if (mode == EmailMode.MailMessage)
                                {
                                    MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 1].Value));
                                    MailMessage message = new MailMessage(from, to);

                                    message.Subject = subject;

                                    message.Attachments.Add(imageInEmailHead);
                                    message.Attachments.Add(imageInEmailImgBlue);
                                    message.Attachments.Add(imageInEmailFooter);

                                    message.IsBodyHtml = true;

                                    message.Body = File.ReadAllText(
                                        AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserPasswordBatchModify.html")
                                        .Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value))
                                        .Replace("${SVNUsername}", Convert.ToString(sheet.Cells[row, 4].Value))
                                        .Replace("${SVNPassword}", Convert.ToString(sheet.Cells[row, 5].Value))
                                        .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                        .Replace("${NEXTDate}", DateTime.Now.AddMonths(1).AddDays(1 - DateTime.Now.Day).ToString("yyyy-MM-dd"))
                                        .Replace("${BEFOREDate}", DateTime.Now.AddMonths(1).AddDays(0 - DateTime.Now.Day).ToString("yyyy-MM-dd"))
                                        .Replace("${AdminEmail}", adminEmail)
                                        .Replace("${AdminName}", adminName)
                                        .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                        );

                                    client.Send(message);
                                }
                            }
                            sheet.Cells[row, 7] = "是";
                            sheet.Cells[row, 8] = DateTime.Now.ToString("yyyy-MM-dd");
                        }
                    }
                    else
                    { // 不需要定期修改密码的
                        if (client != null)
                        {
                            if (mode == EmailMode.MailMessage)
                            {
                                MailAddress to = new MailAddress(Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 1].Value));
                                MailMessage message = new MailMessage(from, to);

                                message.Subject = subject;

                                message.Attachments.Add(imageInEmailHead);
                                message.Attachments.Add(imageInEmailImgBlue);
                                message.Attachments.Add(imageInEmailFooter);

                                message.IsBodyHtml = true;

                                message.Body = File.ReadAllText(
                                    AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/UserPasswordBatchSkiped.html")
                                    .Replace("${Username}", Convert.ToString(sheet.Cells[row, 1].Value))
                                    .Replace("${SVNUsername}", Convert.ToString(sheet.Cells[row, 4].Value))
                                    .Replace("${SVNPassword}", Convert.ToString(sheet.Cells[row, 5].Value))
                                    .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                    .Replace("${NEXTDate}", DateTime.Now.AddMonths(1).AddDays(1 - DateTime.Now.Day).ToString("yyyy-MM-dd"))
                                    .Replace("${BEFOREDate}", DateTime.Now.AddMonths(1).AddDays(0 - DateTime.Now.Day).ToString("yyyy-MM-dd"))
                                    .Replace("${AdminEmail}", adminEmail)
                                    .Replace("${AdminName}", adminName)
                                    .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                    );

                                client.Send(message);
                            }
                        }
                        sheet.Cells[row, 7] = "是";
                        sheet.Cells[row, 8] = DateTime.Now.ToString("yyyy-MM-dd");
                    }
                }
            }
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteUser_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["删除用户"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 6].Value == null || "".Equals(sheet.Cells[row, 6].Value.Trim()))
                {
                    if (sheet.Cells[row, 4].Value == null || "".Equals(sheet.Cells[row, 4].Value))
                    {
                        sheet.Cells[row, 4].Value = Convert.ToString(sheet.Cells[row, 3].Value).Split('@')[0].ToLower().Replace("_", ".");
                    }
                    SVNHelper.DeleteUser(Convert.ToString(sheet.Cells[row, 4].Value));

                    sheet.Cells[row, 6] = "是";
                    sheet.Cells[row, 7] = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }
        }

        #endregion

        #region 权限设置

        /// <summary>
        /// 获取仓库信息
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private Dictionary<string, string> GetRepositoryInfo(string name)
        {
            Dictionary<string, string> repository = new Dictionary<string, string>();


            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["创建仓库"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;

            while (++row <= rows)
            {
                if (sheet.Cells[row, 1].Value == name)
                {
                    repository.Add("RepositoryDesc", sheet.Cells[row, 3].Value);
                    repository.Add("ContactDisplay", sheet.Cells[row, 4].Value);
                    repository.Add("ContactAddress", sheet.Cells[row, 5].Value);
                    break;
                }
            }
            return repository;
        }

        /// <summary>
        /// 获取用户列表
        /// </summary>
        /// <returns></returns>
        private List<Dictionary<string, string>> GetUsersList()
        {
            List<Dictionary<string, string>> results = new List<Dictionary<string, string>>();

            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["创建用户"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 1;

            while (++row <= rows)
            {
                Dictionary<string, string> user = new Dictionary<string, string>();
                user.Add("Username", sheet.Cells[row, 1].Value);
                user.Add("Emailadd", sheet.Cells[row, 3].Value);
                user.Add("SVNUser", sheet.Cells[row, 4].Value);
                results.Add(user);
            }
            return results;
        }

        /// <summary>
        /// 设置仓库人员
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRepositoryMembers_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["通用配置"];

            string manager = sheet.Cells[1, 2].Value;
            string subject = sheet.Cells[2, 2].Value;
            string adminEmail = Convert.ToString(sheet.Cells[5, 2].Value);
            string adminName = Convert.ToString(sheet.Cells[7, 2].Value);
            string repositoryRoot = Convert.ToString(sheet.Cells[8, 2].Value);

            SmtpClient client = null;
            MailAddress from = null;
            try
            {
                client = new SmtpClient(Convert.ToString(sheet.Cells[3, 2].Value), Convert.ToInt32(sheet.Cells[4, 2].Value));
                client.Credentials = new NetworkCredential(adminEmail, Convert.ToString(sheet.Cells[6, 2].Value));
                client.EnableSsl = true;

                from = new MailAddress(adminEmail, adminName);
            }
            catch
            {
                client = null;
            }

            sheet = workbook.Sheets["设置仓库成员"];

            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;

            Attachment imageInEmailHead = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailHead.png");
            imageInEmailHead.ContentId = "imageInEmailHead";

            Attachment imageInEmailImgBlue = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailImgBlue.png");
            imageInEmailImgBlue.ContentId = "imageInEmailImgBlue";

            Attachment imageInEmailFooter = new Attachment(AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/images/imageInEmailFooter.png");
            imageInEmailFooter.ContentId = "imageInEmailFooter";

            StringBuilder builder = new StringBuilder();
            Dictionary<string, Dictionary<string, string>> managers = null;

            List<Dictionary<string, string>> userList = GetUsersList();

            while (++row <= rows)
            {
                if (sheet.Cells[row, 5].Value == null || "".Equals(sheet.Cells[row, 5].Value.Trim()))
                {
                    managers = new Dictionary<string, Dictionary<string, string>>();
                    bool result = false;
                    if (Convert.ToString(sheet.Cells[row, 3].Value) == "添加成员")
                    {
                        result = SVNHelper.SetRepositoryEntryPermission(Convert.ToString(sheet.Cells[row, 1].Value), "/", Convert.ToString(sheet.Cells[row, 2].Value), Convert.ToString(sheet.Cells[row, 4].Value));

                    }
                    else if (Convert.ToString(sheet.Cells[row, 3].Value) == "删除成员")
                    {
                        result = SVNHelper.SetRepositoryEntryPermission(Convert.ToString(sheet.Cells[row, 1].Value), "/", Convert.ToString(sheet.Cells[row, 2].Value));

                    }
                    if (result)
                    {
                        if (client != null)
                        {
                            if (mode == EmailMode.MailMessage)
                            {
                                string emailTemplate = Convert.ToString(sheet.Cells[row, 3].Value) == "添加成员" ? "/email/svn/RepositoryUserAdd.html" : "/email/svn/RepositoryUserRemove.html";

                                string operType = Convert.ToString(sheet.Cells[row, 3].Value).Substring(0, 2);

                                // 获取所有用户
                                List<string> temps = SVNHelper.GetGroupRecursiveUsersName(Convert.ToString(sheet.Cells[row, 2].Value));

                                List<string> users = new List<string>();
                                foreach (string temp in temps)
                                {
                                    if (!users.Contains(temp))
                                    {
                                        users.Add(temp);
                                    }
                                }

                                foreach (string temp in users)
                                {
                                    Dictionary<string, string> user = userList.Find(u => u["SVNUser"].Equals(temp));
                                    if (user != null)
                                    {
                                        string email = "名称：" + user["Username"] + "，&nbsp;邮箱：" + user["Emailadd"] + "，&nbsp;账号：" + temp + "<br />";

                                        if (managers.ContainsKey(Convert.ToString(sheet.Cells[row, 1].Value)))
                                        {
                                            string old = ((Dictionary<string, string>)managers[Convert.ToString(sheet.Cells[row, 1].Value)])["NEWUserlist"];
                                            ((Dictionary<string, string>)managers[Convert.ToString(sheet.Cells[row, 1].Value)])["NEWUserlist"] = old + email;
                                        }
                                        else
                                        {
                                            Dictionary<string, string> repository = GetRepositoryInfo(Convert.ToString(sheet.Cells[row, 1].Value));
                                            repository.Add("NEWUserlist", email);
                                            managers.Add(Convert.ToString(sheet.Cells[row, 1].Value), repository);
                                        }

                                        MailAddress to = new MailAddress(user["Emailadd"], user["Username"]);
                                        MailMessage message = new MailMessage(from, to);

                                        message.Subject = subject;

                                        message.Attachments.Add(imageInEmailHead);
                                        message.Attachments.Add(imageInEmailImgBlue);
                                        message.Attachments.Add(imageInEmailFooter);

                                        message.IsBodyHtml = true;

                                        message.Body = File.ReadAllText(
                                            AppDomain.CurrentDomain.SetupInformation.ApplicationBase + emailTemplate)
                                            .Replace("${Username}", user["Username"])
                                            .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                            .Replace("${SVNRepositoryName}", Convert.ToString(sheet.Cells[row, 1].Value))
                                            .Replace("${SVNRepositoryDesc}", ((Dictionary<string, string>)managers[Convert.ToString(sheet.Cells[row, 1].Value)])["RepositoryDesc"])
                                            .Replace("${AdminEmail}", adminEmail)
                                            .Replace("${AdminName}", adminName)
                                            .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                            );

                                        client.Send(message);

                                    }
                                    else
                                    {
                                        string email = "名称：未知，&nbsp;邮箱：未知，&nbsp;账号：" + temp + "<br />";

                                        if (managers.ContainsKey(Convert.ToString(sheet.Cells[row, 1].Value)))
                                        {
                                            string old = ((Dictionary<string, string>)managers[Convert.ToString(sheet.Cells[row, 1].Value)])["NEWUserlist"];
                                            ((Dictionary<string, string>)managers[Convert.ToString(sheet.Cells[row, 1].Value)])["NEWUserlist"] = old + email;

                                        }
                                        else
                                        {
                                            Dictionary<string, string> repository = GetRepositoryInfo(Convert.ToString(sheet.Cells[row, 1].Value));
                                            repository.Add("NEWUserlist", email);
                                            managers.Add(Convert.ToString(sheet.Cells[row, 1].Value), repository);
                                        }
                                    }
                                }

                                var keys = managers.Keys;
                                foreach (string key in keys)
                                {
                                    var repository = managers[key];
                                    if (repository.Count > 0)
                                    {
                                        IDictionary<string, SVNHelper.AccessLevel> permissions = SVNHelper.GetPermissions(key, "/");

                                        IDictionary<string, SVNHelper.AccessLevel> accountPermission = new Dictionary<string, SVNHelper.AccessLevel>();

                                        foreach (string account in permissions.Keys)
                                        {
                                            temps = SVNHelper.GetGroupRecursiveUsersName(account);
                                            foreach (string name in temps)
                                            {
                                                Dictionary<string, string> user = userList.Find(u => u["SVNUser"].Equals(name));
                                                string email = null;
                                                if (user != null)
                                                {
                                                    email = "名称：" + user["Username"] + "，&nbsp;邮箱：" + user["Emailadd"] + "，&nbsp;账号：" + name + "，&nbsp;权限：";
                                                }
                                                else
                                                {
                                                    email = "名称：未知，&nbsp;邮箱：未知，&nbsp;账号：" + name + "，&nbsp;权限：";
                                                }
                                                if (!accountPermission.ContainsKey(email))
                                                {
                                                    accountPermission[email] = permissions[account];
                                                }
                                                else
                                                {
                                                    accountPermission[email] = accountPermission[email] < permissions[account] ? accountPermission[email] : permissions[account];
                                                }
                                            }
                                        }

                                        builder = new StringBuilder();

                                        foreach (string account in accountPermission.Keys)
                                        {
                                            builder.Append(account).Append(accountPermission[account]).Append("<br />");
                                        }

                                        MailAddress to = new MailAddress(repository["ContactAddress"], repository["ContactDisplay"]);
                                        MailMessage message = new MailMessage(from, to);

                                        message.Bcc.Add(adminEmail);

                                        message.Subject = subject;

                                        message.Attachments.Add(imageInEmailHead);
                                        message.Attachments.Add(imageInEmailImgBlue);
                                        message.Attachments.Add(imageInEmailFooter);

                                        message.IsBodyHtml = true;

                                        message.Body = File.ReadAllText(
                                            AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "/email/svn/RepositoryUserManager.html")
                                            .Replace("${ManagerName}", repository["ContactDisplay"])
                                            .Replace("${OperType}", operType)
                                            .Replace("${NEWUserlist}", repository["NEWUserlist"])
                                            .Replace("${SVNRepositoryRoot}", repositoryRoot)
                                            .Replace("${SVNRepositoryName}", key)
                                            .Replace("${SVNRepositoryDesc}", repository["RepositoryDesc"])
                                            .Replace("${SVNRepositoryPermission}", builder.ToString())
                                            .Replace("${AdminEmail}", adminEmail)
                                            .Replace("${AdminName}", adminName)
                                            .Replace("${SENDDatetime}", DateTime.Now.ToString("yyyy-MM-dd")
                                            );
                                        client.Send(message);
                                    }
                                }
                            }
                        }
                        sheet.Cells[row, 5] = "是";
                        sheet.Cells[row, 6] = DateTime.Now.ToString("yyyy-MM-dd");
                    } // if success
                } // if empty
            } // while
        }

        /// <summary>
        /// 设置仓库条目权限
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRepositoryEntryPermission_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet sheet = workbook.Sheets["设置仓库条目权限"];
            int rows = sheet.UsedRange.Rows.Count;
            int row = 0;
            while (++row <= rows)
            {
                if (sheet.Cells[row, 6].Value == null || "".Equals(sheet.Cells[row, 6].Value.Trim()))
                {
                    if (Convert.ToString(sheet.Cells[row, 4].Value) == "添加成员")
                    {
                        SVNHelper.SetRepositoryEntryPermission(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value), Convert.ToString(sheet.Cells[row, 3].Value), Convert.ToString(sheet.Cells[row, 5].Value));
                    }
                    else if (Convert.ToString(sheet.Cells[row, 4].Value) == "删除成员")
                    {
                        SVNHelper.SetRepositoryEntryPermission(Convert.ToString(sheet.Cells[row, 1].Value), Convert.ToString(sheet.Cells[row, 2].Value), Convert.ToString(sheet.Cells[row, 3].Value));
                    }
                    sheet.Cells[row, 6] = "是";
                    sheet.Cells[row, 7] = DateTime.Now.ToString("yyyy-MM-dd");
                }
            }
        }

        #endregion
    }
}
