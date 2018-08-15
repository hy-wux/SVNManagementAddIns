SVN环境管理
=====
基于Excel的Windows VisualSVN管理工具。  
Windows SVN服务器软件VisualSVN Server提供了图形化的管理界面，但是也存在一些操作上的不方便：
+ 仓库增加目录不能批量增加，只能逐个增加，多级目录不能一次创建
+ Group中添加User没有查找功能，没法快速定位到需要添加的User
+ 仓库、条目给Group或User设置特殊权限的时候没有查找功能，没法快速定位到需要添加的Group和User
+ 没有记录仓库、Group、User的一些附加信息的功能

鉴于以上一些不足之处，考虑到可以利用Excel的方便操作、支持查找、支持筛选，以及能够记录一些自定义的附加信息等特性，因此开发这样一个基于Excel的SVN管理工具。可以实现：
+ 创建仓库、创建仓库目录、重命名仓库、删除仓库
+ 创建组、组中添加用户、组中移除用户、删除组
+ 创建用户、给用户指定组、设置用户密码、删除用户
+ 给仓库指定用户访问权限、给仓库中的目录指定用户访问权限、删除权限
+ 并且结合Excel的特性，可以记录仓库、组、用户的附加信息，再结合程序功能，可以实现每次重要变动操作都可以以邮件的方式通知到相关人员

# 版本历史
## 2018-04-08
1. 支持的功能：
>>+ 创建SVN仓库，并以邮件的方式通知仓库对应的项目负责人；
>>+ 创建SVN仓库的一级目录；
>>+ 重命名仓库；
>>+ 删除仓库；
>>+ 创建用户组；
>>+ 删除用户组；
>>+ 创建用户：根据提供的员工邮箱，生成用户账号、密码，创建用户，并以邮件的方式通知到员工，同时将新建的账号列表通知到管理员；
>>+ 删除用户；
>>+ SVN仓库添加用户：为仓库添加用户，并以邮件的方式通知到用户，同时以邮件的方式将添加的用户列表通知到仓库对应的项目负责人；
>>+ 仓库条目权限控制：仓库中的指定路径，对给定的用户赋予权限，支持NoAccess\ReadOnly\ReadWrite

2. 限制及缺少的功能
>>+ 仅支持Windows平台运行；
>>+ 仅支持Windows平台的VisualSVN Server的管理；
>>+ 仅支持Microsoft Excel 2010及以上版本的Excel；
>>+ 暂不支持创建仓库下的多级目录；
>>+ 暂不支持为用户组添加用户、为用户指定用户组；
>>+ 暂不支持为仓库添加用户组的访问权限；
>>+ 暂不支持从组中移除用户；
>>+ 暂不支持移除用户对SVN仓库的访问权限；
>>+ 暂不支持移除用户对仓库条目的特殊权限设置


## 2018-04-26
1. 新增功能
>>+ 支持为用户组添加用户、为用户指定用户组；
>>+ 支持为仓库添加用户组的访问权限；

2. 限制及缺少的功能
>>+ 仅支持Windows平台运行；
>>+ 仅支持Windows平台的VisualSVN Server的管理；
>>+ 仅支持Microsoft Excel 2010及以上版本的Excel；
>>+ 暂不支持创建仓库下的多级目录；
>>+ 暂不支持从组中移除用户；
>>+ 暂不支持移除用户对SVN仓库的访问权限；
>>+ 暂不支持移除用户对仓库条目的特殊权限设置


## 2018-06-25
1. 新增功能
>>+ 支持创建仓库下的多级目录；

2. 限制及缺少的功能
>>+ 仅支持Windows平台运行；
>>+ 仅支持Windows平台的VisualSVN Server的管理；
>>+ 仅支持Microsoft Excel 2010及以上版本的Excel；
>>+ 暂不支持从组中移除用户；
>>+ 暂不支持移除用户对SVN仓库的访问权限；
>>+ 暂不支持移除用户对仓库条目的特殊权限设置


## 2018-08-03
1. 新增功能
>>+ 支持从组中移除用户；
>>+ 支持移除用户对SVN仓库的访问权限；
>>+ 支持移除用户对仓库条目的特殊权限设置；

2. 限制及缺少的功能
>>+ 仅支持Windows平台运行；
>>+ 仅支持Windows平台的VisualSVN Server的管理；
>>+ 仅支持Microsoft Excel 2010及以上版本的Excel；

# 部分截图
邮件相关的配置：主题、发件服务器、账号、密码等
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/001.png)  

创建一个仓库：仓库名称、用途、相关负责人等
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/002.png)  

SVN服务器会按照配置进行仓库的创建
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/003.png)  

并且会将仓库的相关信息以邮件的方式通知到相关人员
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/004.png)  

创建仓库的目录结构，由于不支持一次创建多级目录，所以需要将上级目录配置在前面，下级目录配置在后面，可以在一行配置或多行配置
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/005.png)  

在仓库中创建了目录
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/006.png)  

创建组
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/007.png)  

成功创建了组
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/008.png)  

创建用户，配置姓名、邮箱，如果配置SVN用户名则以配置的为准，否则默认以邮箱前缀作为SVN用户名
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/009.png)  

创建成功后会发送邮件给相关人员
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/010.png)  

成功创建了用户
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/011.png)  

给仓库赋予人员的访问权限
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/012.png)  

赋予成功
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/013.png)  

以邮件的方式告诉用户有权限访问该仓库
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/014.png)  

以邮件的方式告诉仓库负责人仓库中添加了哪个用户，以及目前该仓库的所有用户及权限
![](https://github.com/hy-wux/SVNManagementAddIns/raw/master/images/015.png)  