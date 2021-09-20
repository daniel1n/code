## 部署说明
> 本项目基于ASP.NET开发的独立站，需要配置相关运行环境

### 1.1.windows安装IIS服务
[win10安装IIS服务 - TechSingularity - 博客园](https://www.cnblogs.com/TechSingularity/p/12017862.html#:~:text=%E4%B8%80%E3%80%81%E4%BD%BF%E7%94%A8%E6%9C%8D%E5%8A%A1%E5%99%A8%E7%AE%A1%E7%90%86%E5%99%A8%E5%AE%89%E8%A3%85IIS%E6%9C%8D%E5%8A%A1%EF%BC%88wind10%EF%BC%89.%201%EF%BC%89%E5%90%AF%E5%8A%A8%E6%8E%A7%E5%88%B6%E9%9D%A2%E6%9D%BF%EF%BC%8C%E9%80%89%E6%8B%A9%E6%9F%A5%E7%9C%8B%E6%96%B9%E5%BC%8F%E4%B8%BA%E2%80%9C%E7%B1%BB%E5%88%AB%E2%80%9D%EF%BC%9B.%202%EF%BC%89%E7%82%B9%E5%87%BB%E2%80%9C%E7%A8%8B%E5%BA%8F%E2%80%9D%EF%BC%9B.%203%EF%BC%89%E7%82%B9%E5%87%BB%E6%89%93%E5%BC%80%E6%88%96%E5%85%B3%E9%97%ADwindow%E5%8A%9F%E8%83%BD%EF%BC%9B.%204%EF%BC%89%E9%80%89%E6%8B%A9Internet%E4%BF%A1%E6%81%AF%E6%9C%8D%E5%8A%A1%EF%BC%8C%E6%A0%87%E7%BA%A2%E6%A1%86%E4%B8%AD%E7%9A%84%E5%85%A8%E9%83%A8%E9%80%89%E4%B8%AD%EF%BC%8C%E5%85%B6%E4%BB%96%E9%80%89%E6%8B%A9%E6%80%A7%E5%8B%BE%E9%80%89%EF%BC%8C%E5%A6%82%E4%B8%8B%E5%9B%BE%E6%89%80%E7%A4%BA%EF%BC%9A.,5%EF%BC%89%E7%82%B9%E5%87%BB%E7%A1%AE%E5%AE%9A.%206%EF%BC%89%E5%9C%A8%E6%B5%8F%E8%A7%88%E5%99%A8%E4%B8%AD%E8%BE%93%E5%85%A5localhost%EF%BC%8C%E5%87%BA%E7%8E%B0%E5%A6%82%E5%9B%BE%E6%89%80%E7%A4%BA%EF%BC%8C%E5%88%99%E9%85%8D%E7%BD%AE%E6%88%90%E5%8A%9F.%207%EF%BC%89%E5%9C%A8%E5%BC%80%E5%A7%8B--%3E%E7%A8%8B%E5%BA%8F%E6%90%9C%E7%B4%A2--%3EIIS%EF%BC%8C%E5%8D%B3%E5%8F%AF%E3%80%82.%20%E4%BA%8C%E3%80%81%E4%BD%BF%E7%94%A8%E5%91%BD%E4%BB%A4%E8%A1%8C%E6%96%B9%E5%BC%8F%E8%BF%9B%E8%A1%8CIIS%E5%AE%89%E8%A3%85%EF%BC%88Windows%20server%202008%EF%BC%89.)

#### 1.1.1.使用服务器管理器安装IIS服务（wind10）

1）启动控制面板，选择查看方式为“类别”；

2）点击“程序”；

3）点击打开或关闭window功能；

4）选择Internet信息服务，标红框中的全部选中，其他选择性勾选，如下图所示：
![avatar](https://img2018.cnblogs.com/i-beta/1758041/201912/1758041-20191210170252383-1738720614.png)


5）点击确定

6）在浏览器中输入localhost，出现如图所示，则配置成功

![avatar](https://img2018.cnblogs.com/i-beta/1758041/201912/1758041-20191210171344339-937422536.png)



 7）在开始-->程序搜索-->IIS，即可。

![avatar](https://img2018.cnblogs.com/i-beta/1758041/201912/1758041-20191210171557473-815689160.png)

 

#### 1.1.2.使用命令行方式进行IIS安装（Windows server 2008）

命令行方式安装，没有用户交互界面，无法获知安装进度，但是可以内嵌在自动化脚本或者程序中在操作系统上安装IIS。

**运行安装命令时需要管理员权限，**具体的安装命令如下：

```shell
@echo off
echo 正在添加IIS功能，这可能需要几分钟时间...
 
start /w pkgmgr /iu:IIS-WebServerRole;IIS-WebServer;IIS-CommonHttpFeatures;IIS-StaticContent;IIS-DefaultDocument;
IIS-DirectoryBrowsing;IIS-HttpErrors;IIS-HttpRedirect;IIS-ApplicationDevelopment;IIS-ASPNET;IIS-NetFxExtensibility;
IIS-ASP;IIS-CGI;IIS-ISAPIExtensions;IIS-ISAPIFilter;IIS-ServerSideIncludes;
IIS-HealthAndDiagnostics;IIS-HttpLogging;IIS-LoggingLibraries;
IIS-RequestMonitor;IIS-HttpTracing;IIS-CustomLogging;
IIS-ODBCLogging;IIS-Security;IIS-BasicAuthentication;
IIS-WindowsAuthentication;IIS-DigestAuthentication;
IIS-ClientCertificateMappingAuthentication;IIS-IISCertificateMappingAuthentication;IIS-URLAuthorization;
IIS-RequestFiltering;IIS-IPSecurity;IIS-Performance;IIS-HttpCompressionStatic;IIS-HttpCompressionDynamic;
IIS-WebServerManagementTools;IIS-ManagementConsole;
IIS-ManagementScriptingTools;IIS-ManagementService;
IIS-IIS6ManagementCompatibility;IIS-Metabase;IIS-WMICompatibility;
IIS-LegacyScripts;IIS-LegacySnapIn;WAS-WindowsActivationService;
WAS-ProcessModel;WAS-NetFxEnvironment;WAS-ConfigurationAPI
 
echo.%errorlevel%
```

start命令用来开启一个新的窗口以执行pkgmgr.exe。/iu参数表示按照指定的名称安装组件，后面跟随的都是IIS中的各种组件名称。

### 1.2.使用NetBox作为部署程序

[NetBox将ASP封装为EXE](https://blog.csdn.net/fadfayger/article/details/5830739)

#### 1.2.1.NetBox是什么

​	**[NetBox](http://www.netbox.cn/)**是一个使用脚本语言进行应用软件开发与发布的开发环境和运行平台，使用 **NetBox**，可以完全使用脚本语言(比如 VBScript，Javascript) 创建出稳定高效的应用软件，并且可以平滑移植到从 Windows 98 到 Windows .NET Server 的全部操作系统上。适用范围对于 WEB 应用，可以迅速将已有的 iis+asp 的应用平滑移植到 **NetBox**应用中，除极少数高级编程外，代码不需要任何修改，同时 **NetBox**还提供大量扩展部件，使得 WEB 应用更加方便。由于 **NetBox**可以将全部代码最终发布成为应用程序，保护了开发人员的利益和代码的完整性。同时， **NetBox**还可以方便地编写更多的桌面应用、系统服务器应用、定制网络应用等等。
​	运行环境要求 **NetBox**的基本运行环境要求很低，最低要求只需要 Windows 98 或者 Windows NT + IE4 即可运行。而如果需要使用系统其他部件(比如 ado)，则需要根据系统情况，如果系统本身未缺省安装，需要自行安装。下面列出的是经过测试的所有系统平台：
​	Windows 98
​	Windows 98 SE
​	Windows ME
​	Windows NT+IE4
​	Windows 2000
​	Windows XP
​	Windows .NET Server

​	以上为该软件的说明文件内的内容。

​	简单的形容就是把ASP文件打包 成一个EXE文件，并且不需要在调试的机器上安装IIS即可正常调试。如果按照说明书来操作的话，观看比较繁琐，本人为方便大家使用，现制作一个简单的使用教程。 

#### 1.2.2.NetBox怎么用

​	[下载NetBox](http://www.netbox.cn/nbsetup.rar)

1. 首先安装NetBox,安装时全部是英文界面，默认安装。

2. 安装完毕后运行桌面上的NetBox Deployment Wizard快捷方式 

3. 打开时有个提示框，是选择软件语言类型的，在此处选择为简体中文，点确定即可。软件界面： 

4.  准备步骤 ：

   1. 安装IIS； 

   2. 在D盘根目录下建立111文件夹（其实在哪个盘符下建立都可以，我是个人喜好。呵呵，您也可以根据自己的喜好变换位置。）； 

   3. 将C盘Inetpub文件夹下的wwwroot文件夹（包含里面9个原始文件）一同拷贝至D盘111文件夹下； 

   4. 把需要封装的ASP文件拷贝至D盘111文件夹下的wwwroot文件夹内（是拷贝至wwwroot文件夹内哦）； 

   5. 在D盘111文件夹内新建 main.box ，将以下内容拷贝进去： 

      ```shell
      Dim httpd
      
      Shell.Service.RunService "NBWeb", "NetBox Web Server", "NetBox Http Server Sample"
      
      '---------------------- Service Event ---------------------
      
      Sub OnServiceStart()
          Set httpd =  NetBox.CreateObject( "NetBox.HttpServer")
      
          If httpd.Create( "", 80) = 0 Then
               Set host = httpd.AddHost( "", "/wwwroot")
      
               host.EnableScript = true
               host.AddDefault "index.asp"
               host.AddDefault "index.asp"
      Shell.Execute """C:/Program Files/Internet Explorer/IEXPLORE.EXE""http://127.0.0.1/"
               httpd.Start
          else
               Shell.Quit 0
          end if
      End Sub
      
      Sub OnServiceStop()
          httpd.Close
      End Sub
      
      Sub OnServicePause()
          httpd.Stop
      End Sub
      
      Sub OnServiceResume()
          httpd.Start
      End Sub
      ```

      代码说明：
      host.AddDefault "default.asp" //首页文件如果为index.asp即更换为index.asp
      host.AddDefault "default.htm" //首页文件如果为index.asp即更换为index.asp 

      

      Shell.Execute """C:/Program Files/Internet Explorer/IEXPLORE.EXE"" http://127.0.0.1/" //这一行是我后加上去的。主要意思是自动使用IE浏览器打开127.0.0.1页面。如果您不想自动打开，您也可以去掉。


      If httpd.Create("", 80) = 0 Then  //80是指80端口 不推荐更改。
           Set host = httpd.AddHost( "","/wwwroot") //wwwroot是指111文件夹下wwwroot文件夹名称 

5. 开始封装：
   1.  打开桌面上的NetBox Deployment Wizard快捷方式； 
   2.  单击选择文件夹选中D盘下的111文件夹； 
   3.  单击浏览选择输出文件保存名称及路径（创建NetBox2.exe文件），之后直接点击Build即可自动生成EXE文件；   
   4. 生成后即可运行，安装过IIS的朋友如果使用的是80端口的话要记得在运行程序之前要现停止IIS服务器才可以运行生成的EXE程序 

6. 预览效果

   ![1632152829850](../code/首页图片.png)

## 文件目录
## 其他说明