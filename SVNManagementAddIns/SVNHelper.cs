using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;

namespace SVNManagementAddIn
{
    class SVNHelper
    {
        #region 仓库管理

        /// <summary>
        /// 创建仓库
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool CreateRepository(string name)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_Repository", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Create"); //创建方法参数引用

                @params["Name"] = name.Trim(); //传入参数

                svn.InvokeMethod("Create", @params, null); //执行
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 创建仓库目录
        /// </summary>
        /// <param name="repositories"> </param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool CreateRepositoryFolders(string repositories, string[] name)
        {
            try
            {
                var repository = new ManagementClass("root\\VisualSVN", "VisualSVN_Repository", null);
                ManagementObject repoObject = repository.CreateInstance();
                if (repoObject != null)
                {
                    repoObject.SetPropertyValue("Name", repositories);
                    ManagementBaseObject inParams = repository.GetMethodParameters("CreateFolders");
                    inParams["Folders"] = name;
                    inParams["Message"] = "";
                    repoObject.InvokeMethod("CreateFolders", inParams, null);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 根据仓库名取得仓库实体
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static ManagementObject GetRepositoryObject(string name)
        {
            return new ManagementObject("root\\VisualSVN", string.Format("VisualSVN_Repository.Name='{0}'", name), null);
        }

        /// <summary>
        /// 重命名仓库
        /// </summary>
        /// <param name="oldname"></param>
        /// <param name="newname"></param>
        /// <returns></returns>
        public static bool RenameRepository(string oldname, string newname)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_Repository", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Rename"); //创建方法参数引用

                @params["OldName"] = oldname.Trim();//传入参数
                @params["NewName"] = newname.Trim();//传入参数

                svn.InvokeMethod("Rename", @params, null); //执行
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 删除仓库
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool DeleteRepository(string name)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_Repository", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Delete"); //创建方法参数引用

                @params["Name"] = name.Trim();//传入参数

                svn.InvokeMethod("Delete", @params, null); //执行
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region 组管理

        /// <summary>
        /// 创建组
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool CreatGroup(string name)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Create");

                @params["Name"] = name.Trim();
                @params["Members"] = new object[] { };

                svn.InvokeMethod("Create", @params, null);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 读取指定组里的成员
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static List<ManagementBaseObject> GetGroupMembersObject(string name)
        {
            var listMembers = new List<ManagementBaseObject>();
            var group = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);
            ManagementObject instance = group.CreateInstance();
            if (instance != null)
            {
                instance.SetPropertyValue("Name", name.Trim());
                ManagementBaseObject outParams = instance.InvokeMethod("GetMembers", null, null);
                if (outParams != null)
                {
                    var members = outParams["Members"] as ManagementBaseObject[];
                    if (members != null)
                    {
                        foreach (ManagementBaseObject member in members)
                        {
                            listMembers.Add(member);
                        }
                    }
                }
            }
            return listMembers;
        }

        /// <summary>
        /// 读取指定组里的成员名称
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static List<string> GetGroupMembersName(string name)
        {
            List<string> listMembers = new List<string>();
            ManagementClass group = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);
            ManagementObject instance = group.CreateInstance();
            if (instance != null)
            {
                instance.SetPropertyValue("Name", name.Trim());
                ManagementBaseObject outParams = instance.InvokeMethod("GetMembers", null, null); //通过实例来调用方法
                if (outParams != null)
                {
                    var members = outParams["Members"] as ManagementBaseObject[];
                    if (members != null)
                    {
                        foreach (ManagementBaseObject member in members)
                        {
                            listMembers.Add(member["Name"].ToString());
                        }
                    }
                }
            }
            return listMembers;
        }

        /// <summary>
        /// 递归获取仓库下的所有人员
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public static List<string> GetGroupRecursiveUsersName(string groupName)
        {
            List<string> results = new List<string>();
            string[] groups = groupName.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string group in groups)
            {
                try
                {
                    List<ManagementBaseObject> groupMembers = GetGroupMembersObject(group);
                    ManagementClass groupClass = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);

                    // 组，获取所有组员
                    foreach (ManagementBaseObject member in groupMembers)
                    {
                        if (member.ClassPath.ClassName.Equals("VisualSVN_Group"))
                        {
                            results.AddRange(GetGroupRecursiveUsersName(member["Name"].ToString()));
                        }
                        else
                        {
                            results.Add(member["Name"].ToString());
                        }
                    }
                }
                catch
                {
                    // 用户，直接添加名称
                    results.Add(group);
                }

            }
            return results;
        }

        /// <summary>
        /// 设置组拥有的用户
        /// </summary>
        /// <param name="groupName"></param>
        /// <param name="userNames"></param>
        /// <param name="operTypes"></param>
        /// <returns></returns>
        public static bool SetGroupMembers(string groupName, string userNames, string operTypes)
        {
            try
            {
                List<string> listMembersName = GetGroupMembersName(groupName);
                List<ManagementBaseObject> listMembers = new List<ManagementBaseObject>();

                string[] names = userNames.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                if (operTypes == "添加成员")
                {
                    foreach (string name in names)
                    {
                        if (!listMembersName.Contains(name))
                            listMembersName.Add(name);
                    }
                }
                else if (operTypes == "删除成员")
                {
                    foreach (string name in names)
                    {
                        listMembersName.Remove(name);
                    }
                }

                foreach (string name in listMembersName)
                {
                    ManagementObject account;
                    try
                    {
                        // 先判断是否是组
                        account = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null).CreateInstance();
                        GetGroupMembersName(name);
                    }
                    catch
                    {
                        // 如果不是组就判断是用户
                        account = new ManagementClass("root\\VisualSVN", "VisualSVN_User", null).CreateInstance();
                    }
                    if (account != null)
                    {
                        account.SetPropertyValue("Name", name);
                        listMembers.Add(account as ManagementBaseObject);
                    }
                }
                var group = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);
                ManagementObject instance = group.CreateInstance();
                if (instance != null)
                {
                    instance.SetPropertyValue("Name", groupName.Trim());

                    ManagementBaseObject @params = instance.GetMethodParameters("SetMembers");

                    @params["Members"] = listMembers.ToArray();

                    instance.InvokeMethod("SetMembers", @params, null);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 删除组
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static bool DeleteGroup(string name)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_Group", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Delete");

                @params["Name"] = name.Trim();

                svn.InvokeMethod("Delete", @params, null);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region 用户管理

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static bool CreateUser(string username, string password)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_User", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Create");

                @params["Name"] = username.Trim();
                @params["Password"] = password.Trim();

                svn.InvokeMethod("Create", @params, null);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 设置用户所属的组
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="groupNames"></param>
        /// <param name="operTypes"></param>
        /// <returns></returns>
        public static bool SetMemberGroup(string userName, string groupNames, string operTypes)
        {
            string[] groups = groupNames.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string group in groups)
            {
                SetGroupMembers(group, userName, operTypes);
            }
            return true;
        }

        /// <summary>
        /// 设置用户密码
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static bool SetPassword(string username, string password)
        {
            try
            {
                var user = new ManagementClass("root\\VisualSVN", "VisualSVN_User", null);
                ManagementObject instance = user.CreateInstance();
                if (instance != null)
                {
                    instance.SetPropertyValue("Name", username.Trim());
                    ManagementBaseObject @params = instance.GetMethodParameters("SetPassword");

                    @params["Password"] = password.Trim();

                    instance.InvokeMethod("SetPassword", @params, null);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        public static bool DeleteUser(string username)
        {
            try
            {
                var svn = new ManagementClass("root\\VisualSVN", "VisualSVN_User", null);
                ManagementBaseObject @params = svn.GetMethodParameters("Delete");

                @params["Name"] = username.Trim();

                svn.InvokeMethod("Delete", @params, null);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region 设置仓库权限

        /// <summary>
        /// 权限列表
        /// </summary>
        public enum AccessLevel : uint
        {
            NoAccess = 0,
            ReadOnly = 1,
            ReadWrite = 2
        }

        /// <summary>
        /// 设置仓库条目权限(添加成员)
        /// </summary>
        /// <param name="repository">仓库名称</param>
        /// <param name="path">仓库条目路径</param>
        /// <param name="name">实体名称</param>
        /// <param name="permission">访问权限</param>
        /// <returns></returns>
        public static bool SetRepositoryEntryPermission(string repository, string path, string name, string permission)
        {
            try
            {
                string[] names = name.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                // 原有的权限
                IDictionary<string, AccessLevel> permissions = GetPermissions(repository, path);
                foreach (string s in names)
                {
                    if (permissions.ContainsKey(s))
                    {
                        // 先删除原有的权限
                        // 目的是支持权限更新
                        permissions.Remove(s);
                    }

                    if (permission.Equals("NoAccess"))
                    {
                        permissions.Add(s, AccessLevel.NoAccess);
                    }
                    else if (permission.Equals("ReadOnly"))
                    {
                        permissions.Add(s, AccessLevel.ReadOnly);
                    }
                    else if (permission.Equals("ReadWrite"))
                    {
                        permissions.Add(s, AccessLevel.ReadWrite);
                    }
                    else
                    {
                        permissions.Add(s, AccessLevel.ReadOnly);
                    }
                }

                SetPermissions(repository, path, permissions);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 设置仓库条目权限(删除成员)
        /// </summary>
        /// <param name="repository">仓库名称</param>
        /// <param name="path">仓库条目路径</param>
        /// <param name="name">实体名称</param>
        /// <returns></returns>
        public static bool SetRepositoryEntryPermission(string repository, string path, string name)
        {
            try
            {
                string[] names = name.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);

                // 原有的权限
                IDictionary<string, AccessLevel> permissions = GetPermissions(repository, path);
                foreach (string s in names)
                {
                    if (permissions.ContainsKey(s))
                    {
                        // 删除原有的权限
                        permissions.Remove(s);
                    }
                }

                SetPermissions(repository, path, permissions);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 读取权限实体
        /// </summary>
        /// <param name="name">用户/组名称</param>
        /// <param name="accessLevel">访问权限</param>
        /// <returns></returns>
        private static ManagementObject GetPermissionObject(string name, AccessLevel accessLevel)
        {
            var entryClass = new ManagementClass("root\\VisualSVN",
                                                 "VisualSVN_PermissionEntry", null);
            ManagementObject entry = entryClass.CreateInstance();
            ManagementClass accountClass;
            if (entry != null)
            {
                try
                { // 先根据组进行权限设置
                    List<ManagementBaseObject> list = GetGroupMembersObject(name);
                    accountClass = new ManagementClass("root\\VisualSVN",
                                                               "VisualSVN_Group", null);
                }
                catch (Exception ex)
                { // 组不存在，就找用户
                    accountClass = new ManagementClass("root\\VisualSVN",
                                                               "VisualSVN_User", null);
                }
                ManagementObject account = accountClass.CreateInstance();
                if (account != null) account["Name"] = name;
                entry["Account"] = account;
                entry["AccessLevel"] = accessLevel;
            }
            return entry;
        }

        /// <summary>
        /// 设置仓库权限
        /// </summary>
        /// <param name="repositoryName">仓库名称</param>
        /// <param name="path">仓库条目路径</param>
        /// <param name="permissions">权限</param>
        private static void SetPermissions(string repositoryName, string path,
                                           IEnumerable<KeyValuePair<string, AccessLevel>> permissions)
        {
            ManagementObject repository = GetRepositoryObject(repositoryName);
            ManagementBaseObject inParameters = repository.GetMethodParameters("SetSecurity");
            inParameters["Path"] = path;
            IEnumerable<ManagementObject> permissionObjects =
                permissions.Select(p => GetPermissionObject(p.Key, p.Value));
            ManagementObject[] objs = permissionObjects.ToArray();
            inParameters["Permissions"] = objs;
            repository.InvokeMethod("SetSecurity", inParameters, null);
        }

        /// <summary>
        /// 读取仓库权限
        /// </summary>
        /// <param name="repositoryName">仓库名称</param>
        /// <param name="path">仓库条目路径</param>
        /// <returns></returns>
        public static IDictionary<string, AccessLevel> GetPermissions(string repositoryName, string path)
        {
            ManagementObject repository = GetRepositoryObject(repositoryName);
            ManagementBaseObject inParameters = repository.GetMethodParameters("GetSecurity");
            inParameters["Path"] = path;
            ManagementBaseObject outParameters = repository.InvokeMethod("GetSecurity", inParameters, null);

            var permissions = new Dictionary<string, AccessLevel>();

            if (outParameters != null)
                foreach (ManagementBaseObject p in (ManagementBaseObject[])outParameters["Permissions"])
                {
                    var account = (ManagementBaseObject)p["Account"];
                    var name = (string)account["Name"];
                    var accessLevel = (AccessLevel)p["AccessLevel"];

                    permissions[name] = accessLevel;
                }

            return permissions;
        }

        #endregion
    }
}
