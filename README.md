# RF

## 环境搭建(仅需一次)
* 下载 Anaconda，地址 https://mirrors.tuna.tsinghua.edu.cn/anaconda/archive/Anaconda2-5.0.1-Windows-x86_64.exe (64位系统) 或 https://mirrors.tuna.tsinghua.edu.cn/anaconda/archive/Anaconda2-5.0.1-Windows-x86.exe (32位系统)
* 双击安装，安装路径可改
* 安装wxPython包
  * 在程序中找到"Anaconda Prompt"，点击运行
  * 输入`conda install -c anaconda wxpython`，回车。若有提示，输入y。（此步骤需要网络）
  * 安装成功，关闭退出
  * PS: 如果下载非常慢，请在执行`conda install ...`之前额外输入以下命令
    * `conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/main/`
    * `conda config --add channels https://mirrors.tuna.tsinghua.edu.cn/anaconda/pkgs/free/`
    * `conda config --set show_channel_urls yes`
## 运行程序
* 下载本工程: https://github.com/wangxfdu/RF/archive/master.zip 解压后会释放main.py文件，如果程序有更新，重新下载即可
* 在程序中找到"Anaconda Prompt"，点击运行
* 输入 `cd c:\path_to_main_py` _注：c:\path_to_main_py为main.py所在文件夹_
* 输入 `python main.py`
