namespace SVNManagementAddIn
{
    partial class ExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExcelRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.TabSVN = this.Factory.CreateRibbonTab();
            this.SVNGroupServer = this.Factory.CreateRibbonGroup();
            this.SVNGroupRight = this.Factory.CreateRibbonGroup();
            this.menuRepositories = this.Factory.CreateRibbonMenu();
            this.btnCreateRepositories = this.Factory.CreateRibbonButton();
            this.btnCreateRepositoriesFolders = this.Factory.CreateRibbonButton();
            this.btnRenameRepositories = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnDeleteRepositories = this.Factory.CreateRibbonButton();
            this.menuGroups = this.Factory.CreateRibbonMenu();
            this.btnCreateGroups = this.Factory.CreateRibbonButton();
            this.btnSetMembers = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnDeleteGroups = this.Factory.CreateRibbonButton();
            this.menuUsers = this.Factory.CreateRibbonMenu();
            this.btnCreateUsers = this.Factory.CreateRibbonButton();
            this.btnSetGroups = this.Factory.CreateRibbonButton();
            this.btnSetPassword = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnEmailVerify = this.Factory.CreateRibbonButton();
            this.btnResetPassword = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.btnDeleteUser = this.Factory.CreateRibbonButton();
            this.btnRepositoryMembers = this.Factory.CreateRibbonButton();
            this.btnRepositoryEntryPermission = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.TabSVN.SuspendLayout();
            this.SVNGroupServer.SuspendLayout();
            this.SVNGroupRight.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // TabSVN
            // 
            this.TabSVN.Groups.Add(this.SVNGroupServer);
            this.TabSVN.Groups.Add(this.SVNGroupRight);
            this.TabSVN.Label = "伍鲜-SVN管理";
            this.TabSVN.Name = "TabSVN";
            // 
            // SVNGroupServer
            // 
            this.SVNGroupServer.Items.Add(this.menuRepositories);
            this.SVNGroupServer.Items.Add(this.menuGroups);
            this.SVNGroupServer.Items.Add(this.menuUsers);
            this.SVNGroupServer.Label = "SVN服务管理";
            this.SVNGroupServer.Name = "SVNGroupServer";
            // 
            // SVNGroupRight
            // 
            this.SVNGroupRight.Items.Add(this.btnRepositoryMembers);
            this.SVNGroupRight.Items.Add(this.btnRepositoryEntryPermission);
            this.SVNGroupRight.Label = "SVN权限管理";
            this.SVNGroupRight.Name = "SVNGroupRight";
            // 
            // menuRepositories
            // 
            this.menuRepositories.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuRepositories.Image = global::SVNManagementAddIn.Properties.Resources.DatabaseCopyDatabaseFile;
            this.menuRepositories.Items.Add(this.btnCreateRepositories);
            this.menuRepositories.Items.Add(this.btnCreateRepositoriesFolders);
            this.menuRepositories.Items.Add(this.btnRenameRepositories);
            this.menuRepositories.Items.Add(this.separator1);
            this.menuRepositories.Items.Add(this.btnDeleteRepositories);
            this.menuRepositories.Label = "仓库管理";
            this.menuRepositories.Name = "menuRepositories";
            this.menuRepositories.ShowImage = true;
            // 
            // btnCreateRepositories
            // 
            this.btnCreateRepositories.Label = "创建仓库";
            this.btnCreateRepositories.Name = "btnCreateRepositories";
            this.btnCreateRepositories.ShowImage = true;
            this.btnCreateRepositories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateRepositories_Click);
            // 
            // btnCreateRepositoriesFolders
            // 
            this.btnCreateRepositoriesFolders.Label = "创建目录";
            this.btnCreateRepositoriesFolders.Name = "btnCreateRepositoriesFolders";
            this.btnCreateRepositoriesFolders.ShowImage = true;
            this.btnCreateRepositoriesFolders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateRepositoriesFolders_Click);
            // 
            // btnRenameRepositories
            // 
            this.btnRenameRepositories.Label = "重命名仓库";
            this.btnRenameRepositories.Name = "btnRenameRepositories";
            this.btnRenameRepositories.ShowImage = true;
            this.btnRenameRepositories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRenameRepositories_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnDeleteRepositories
            // 
            this.btnDeleteRepositories.Label = "删除仓库";
            this.btnDeleteRepositories.Name = "btnDeleteRepositories";
            this.btnDeleteRepositories.ShowImage = true;
            this.btnDeleteRepositories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteRepositories_Click);
            // 
            // menuGroups
            // 
            this.menuGroups.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuGroups.Image = global::SVNManagementAddIn.Properties.Resources.GroupMembers;
            this.menuGroups.Items.Add(this.btnCreateGroups);
            this.menuGroups.Items.Add(this.btnSetMembers);
            this.menuGroups.Items.Add(this.separator2);
            this.menuGroups.Items.Add(this.btnDeleteGroups);
            this.menuGroups.Label = "组管理";
            this.menuGroups.Name = "menuGroups";
            this.menuGroups.ShowImage = true;
            // 
            // btnCreateGroups
            // 
            this.btnCreateGroups.Label = "创建组";
            this.btnCreateGroups.Name = "btnCreateGroups";
            this.btnCreateGroups.ShowImage = true;
            this.btnCreateGroups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateGroups_Click);
            // 
            // btnSetMembers
            // 
            this.btnSetMembers.Label = "设置组员";
            this.btnSetMembers.Name = "btnSetMembers";
            this.btnSetMembers.ShowImage = true;
            this.btnSetMembers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetMembers_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnDeleteGroups
            // 
            this.btnDeleteGroups.Label = "删除组";
            this.btnDeleteGroups.Name = "btnDeleteGroups";
            this.btnDeleteGroups.ShowImage = true;
            this.btnDeleteGroups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteGroups_Click);
            // 
            // menuUsers
            // 
            this.menuUsers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuUsers.Image = global::SVNManagementAddIn.Properties.Resources.InviteAttendees;
            this.menuUsers.Items.Add(this.btnCreateUsers);
            this.menuUsers.Items.Add(this.btnSetGroups);
            this.menuUsers.Items.Add(this.btnSetPassword);
            this.menuUsers.Items.Add(this.separator3);
            this.menuUsers.Items.Add(this.btnEmailVerify);
            this.menuUsers.Items.Add(this.btnResetPassword);
            this.menuUsers.Items.Add(this.separator4);
            this.menuUsers.Items.Add(this.btnDeleteUser);
            this.menuUsers.Label = "用户管理";
            this.menuUsers.Name = "menuUsers";
            this.menuUsers.ShowImage = true;
            // 
            // btnCreateUsers
            // 
            this.btnCreateUsers.Label = "创建用户";
            this.btnCreateUsers.Name = "btnCreateUsers";
            this.btnCreateUsers.ShowImage = true;
            this.btnCreateUsers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateUsers_Click);
            // 
            // btnSetGroups
            // 
            this.btnSetGroups.Label = "设置组别";
            this.btnSetGroups.Name = "btnSetGroups";
            this.btnSetGroups.ShowImage = true;
            this.btnSetGroups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetGroups_Click);
            // 
            // btnSetPassword
            // 
            this.btnSetPassword.Label = "设置密码";
            this.btnSetPassword.Name = "btnSetPassword";
            this.btnSetPassword.ShowImage = true;
            this.btnSetPassword.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetPassword_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btnEmailVerify
            // 
            this.btnEmailVerify.Label = "邮箱验证";
            this.btnEmailVerify.Name = "btnEmailVerify";
            this.btnEmailVerify.ShowImage = true;
            this.btnEmailVerify.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmailVerify_Click);
            // 
            // btnResetPassword
            // 
            this.btnResetPassword.Label = "重置密码";
            this.btnResetPassword.Name = "btnResetPassword";
            this.btnResetPassword.ShowImage = true;
            this.btnResetPassword.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnResetPassword_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // btnDeleteUser
            // 
            this.btnDeleteUser.Label = "删除用户";
            this.btnDeleteUser.Name = "btnDeleteUser";
            this.btnDeleteUser.ShowImage = true;
            this.btnDeleteUser.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteUser_Click);
            // 
            // btnRepositoryMembers
            // 
            this.btnRepositoryMembers.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRepositoryMembers.Image = global::SVNManagementAddIn.Properties.Resources.GroupAdminister;
            this.btnRepositoryMembers.Label = "设置仓库成员";
            this.btnRepositoryMembers.Name = "btnRepositoryMembers";
            this.btnRepositoryMembers.ShowImage = true;
            this.btnRepositoryMembers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRepositoryMembers_Click);
            // 
            // btnRepositoryEntryPermission
            // 
            this.btnRepositoryEntryPermission.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRepositoryEntryPermission.Image = global::SVNManagementAddIn.Properties.Resources.NewContact;
            this.btnRepositoryEntryPermission.Label = "设置仓库条目权限";
            this.btnRepositoryEntryPermission.Name = "btnRepositoryEntryPermission";
            this.btnRepositoryEntryPermission.ShowImage = true;
            this.btnRepositoryEntryPermission.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRepositoryEntryPermission_Click);
            // 
            // ExcelRibbon
            // 
            this.Name = "ExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.TabSVN);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.TabSVN.ResumeLayout(false);
            this.TabSVN.PerformLayout();
            this.SVNGroupServer.ResumeLayout(false);
            this.SVNGroupServer.PerformLayout();
            this.SVNGroupRight.ResumeLayout(false);
            this.SVNGroupRight.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab TabSVN;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SVNGroupServer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateUsers;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SVNGroupRight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateGroups;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateRepositories;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuRepositories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRenameRepositories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteRepositories;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateRepositoriesFolders;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuGroups;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteGroups;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetMembers;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuUsers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetGroups;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetPassword;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRepositoryEntryPermission;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRepositoryMembers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnResetPassword;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEmailVerify;
    }

    partial class ThisRibbonCollection
    {
        internal ExcelRibbon Ribbon
        {
            get { return this.GetRibbon<ExcelRibbon>(); }
        }
    }
}
