# ExcelHandler安装部署

## 简介  

ExcelHandler应用程序通过uwsgi接口提供Excel文件预处理，程序由两部分组成：  

- main_prog.py ——主程序，负责接口提供、文件管理等。
- excel_handler.py ——功能库，包括各种excel文件处理函数和方法。

## 环境

- Python 3.5+ (xlrd, openpyxl, Pandas, SciPy, NumPy, Flask)  
   程序是用Python3写的，用到的外部库包括xlrd(读取.xls文件)、openpyxl(读写.xlsx文件)、Pandas(数据处理)、SciPy(科学计算和机器学习)、NumPy(代数运算)、Flask(web服务框架)
- uWSGI  
   通过uwsgi接口提供服务的服务器
- systemd(CentOS) 或 Supervisor(Debian、Ubuntu)  
   维持服务的守护进程

## 安装及环境搭建

安装过程使用root用户进行：  

> [user@server]$ sudo -s  

默认语言编码<u>**必须**</u>是UTF-8系列类型，因此，必须检查`/etc/locale.conf`文件有以下描述：  

> LANG="en_US.UTF-8"  
> LC_ALL="en_US.UTF-8"  
> LC_LANG="en_US.UTF-8"  

若修改了系统默认语言编码需重启系统。

### （一） Python3

#### 1.1 Python3

大部分Linux发行版都没有集成最新的Python3或者编译版本的安装源，考虑到使用到的外部库比较多，从源码编译Python3环境和外部库比较麻烦，考虑选择Anaconda的Python发行版，选择只包含python和conda的Miniconda。  

1) 用清华源下载installer文件：  

> [root@server]# wget https://mirrors.tuna.tsinghua.edu.cn/anaconda/miniconda/Miniconda3-latest-Linux-x86\_64.sh  

2) 开始安装：  

> [root@server]# bash Miniconda3-latest-Linux-x86\_64.sh  

3) 安装过程中，修改默认的安装路径`/root/miniconda3`为`/usr/anaconda3`  

> Miniconda3 will now be installed into this location:  
> /root/miniconda3  
> \- Press ENTER to confirm the location  
> \- Press CTRL-C to abort the installation  
> \- Or specify a different location below  
> [/root/miniconda3] >>> /usr/anaconda3  

4) 选择写入路径至PATH环境参数  

> Do you wish the installer to prepend the Miniconda3 install location to PATH in your /root/.bashrc ? [yes|no]  
> [no] >>> yes  

5) 安装完毕，检查一下路径是否被写进$PATH参数：  

> [root@server]# source ~/.bashrc  
> [root@server]# whereis python3  
> python3: /usr/anaconda3/bin/python3 /usr/anaconda3/bin/python3.6 /usr/anaconda3/bin/python3.6-config /usr/anaconda3/bin/python3.6m /usr/anaconda3/bin/python3.6m-config  

6) 为conda添加清华源并更新全部组件：  

> [root@server]# conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/  
> [root@server]# conda config --set show_channel_urls yes  
> [root@server]# conda update --all  

至此Miniconda3的Pyhton3基础环境安装完毕。  

#### 1.2 NumPy, SciPy, Pandas, xlrd, openpyxl, Flask, pcre

通过conda安装需要的Python库到conda的根环境（若使用虚拟环境，需在相应虚拟环境中再次通过conda安装）。  

> [root@server]# conda install numpy scipy pandas xlrd openpyxl flask pcre

至此，Python3环境已经就绪。  

### （二） uWSGI server

conda没有提供uWSGI服务器，虽然可以用Python的pip安装，但用源码安装方可保证正常稳定运行。  
#### 2.1 源码下载  

在[uWSGI的官方文档网页](http://uwsgi-docs.readthedocs.io/en/latest/Download.html)可以找到最新版本的源码链接。  
下载官方源码，解压后移到`/usr`目录中：  

> [root@server]# wget https://projects.unbit.it/downloads/uwsgi-2.0.15.tar.gz  
> [root@server]# tar -xvzf uwsgi-2.0.15.tar.gz  
> [root@server]# mv uwsgi-2.0.15 /usr  

#### 2.2 编译安装

确认系统默认的python命令使用Python3版本后（用`python -V`命令确认），进入`/usr/uwsgi-2.0.15`目录，编译安装uWSGI：  

> [root@server]# python uwsgiconfig.py --build  
> :  
> :  
> \######## end of uWSGI configuration ########  
> total build time: 23 seconds  
> \*\*\* uWSGI is ready, launch it with ./uwsgi \*\*\*  

至此uWSGI server安装完毕。测试是否可以正常运行：  

>   [root@server]# uwsgi  
>   \*\*\* Starting uWSGI 2.0.15 (64bit) on [Tue Jun 13 13:31:05 2017] \*\*\*  
>   compiled with version: 4.8.5 20150623 (Red Hat 4.8.5-11) on 19 May 2017 14:33:49  
>   os: Linux-3.10.0-514.21.1.el7.x86_64 #1 SMP Thu May 25 17:04:51 UTC 2017  
>   nodename: etown  
>   machine: x86_64  
>   clock source: unix  
>   pcre jit disabled  
>   detected number of CPU cores: 4  
>   current working directory: /usr/bin  
>   detected binary path: /usr/sbin/uwsgi  
>   uWSGI running as root, you can use --uid/--gid/--chroot options  
>   \*\*\* WARNING: you are running uWSGI as root !!! (use the --uid flag) \*\*\*  
>   \*\*\* WARNING: you are running uWSGI without its master process manager \*\*\*  
>   your processes number limit is 14356  
>   your memory page size is 4096 bytes  
>   detected max file descriptor number: 1024  
>   lock engine: pthread robust mutexes  
>   thunder lock: disabled (you can enable it with --thunder-lock)  
>   The -s/--socket option is missing and stdin is not a socket.  

#### 2.3 至关重要的一步！！！

uWSGI默认会从`/usr/sbin/uwsgi`运行，但从systemd启动服务必须在`/usr/anaconda3/bin`运行，因此需要建立一个软链接：  

>   [root@server]# ln -s /usr/uwsgi-2.0.15/uwsgi /usr/anaconda3/bin/uwsgi  

**<u>完成这一步应用服务才能正常地跑起来！</u>**

### （三） ExcelHandler安装及环境配置  

#### 3.1 程序文件  

应用程序包括以下部分：  

-   main_prog.py ——主程序，负责接口提供、文件管理等。
-   excel_handler.py ——功能库，包括各种excel文件处理函数和方法。
-   ehApp.ini ——uwsgi服务启动参数文件。
-   ehApp.service ——用于添加systemd服务的文件

#### 3.2 程序目录结构  

工作目录必须在`/home/webApp/ehApp`，并且具有以下目录结构：  

>home  
>'   └─ webApp  
>'               └─ ehApp (程序根目录)   
>'			├─ templates (页面模板目录（测试用）)  
>'			├─ infiles  (传入文件临时保存的目录)  
>'			└─ outfiles (处理完成输出文件的目录)  

需手动建立`/home/webApp`目录，将`ehApp.tar`文件解压到该目录上：  
>   [root@server]# mkdir /home/webApp  
>   [root@server]# cd /home/webApp  
>   [root@server]# cp /保存ehApp程序压缩包的路径/path/to/ehApp.tar ./  
>   [root@server]# tar -xvf ehApp.tar  
>   [root@server]# chmod -R 755 ./ehApp  

至此，ExcelHandler程序已经安装就绪。  

#### 3.3 ExcelHandler服务配置说明  

ExcelHandler的服务启动参数位于`ehApp.ini`文件，配置说明详见文件注释。默认配置为只可以通过本地的9999端口（即`127.0.0.1:9999`）访问服务，如需全网访问，修改`http-socket`配置项为`0.0.0.0:9999`

## 部署系统服务  

CentOS系统用systemd管理服务的自启动和服务进程守护，设置ExcelHandler程序为随系统启动的服务进程，需要先将ehApp.service文件拷贝至systemd的指定目录并配置激活服务：  
>[root@server]# cp /home/webApp/ehApp/ehApp.service /etc/systemd/system  
>[root@server]# systemctl enable ehApp.service  
>[root@server]# systemctl start ehApp.service  

检查服务启动状态，正常应该显示类似如下的信息：
>[root@server]# systemctl status ehApp.service  
>● ehApp.service - Excel Handler application by uWSGI  
>Loaded: loaded (/etc/systemd/system/ehApp.service; enabled; vendor preset: disabled)  
>Active: active (running) since 二 2017-06-13 16:59:25 CST; 2 days ago  
>Main PID: 22377 (uwsgi)  
>Status: "uWSGI is ready"  
>CGroup: /system.slice/ehApp.service  
>'			├─22377 /usr/anaconda3/bin/uwsgi --ini /home/webApp/ehApp/ehApp.ini  
>'			└─22387 /usr/anaconda3/bin/uwsgi --ini /home/webApp/ehApp/ehApp.ini  
>16:59:25 server  uwsgi[22377]: WSGI app 0 (mountpoint='') ready in 3 seconds on interpreter 0x1c64f30 pid: 22377 (default app)   

用浏览器或者`curl`访问`127.0.0.1:9999`，可以得到程序工作目录的json返回。