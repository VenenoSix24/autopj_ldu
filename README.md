# autopj_ldu
# 鲁东大学教务自动评教脚本 

## 食用方法

1. 下载 [Chrome](https://www.google.com/chrome/)

   > 如果你已安装，请修改 run.py 里的 **CHROME_EXECUTABLE_PATH** ，指定你的 Chrome 路径

2. 下载 符合你 Chrome 版本的 [ChromeDriver](https://googlechromelabs.github.io/chrome-for-testing/) 并将其放到与 ***run.py*** 同级目录下

3. Python 版本 >= **3.11**

4. 安装依赖

   `pip install -r requirements.txt 或者 pip install selenium`

5. 运行 ***run.py***

   创建虚拟环境(可选) 

   ​	`python3 -m venv venv` 运行 `/autopj_ldu/venv/Scripts/Activate.ps1` 启动虚拟环境

   ​	退出虚拟环境 `deactivate` 

   运行脚本

   `python run.py`

6. 按照提示脚本 登录鲁东大学教务系统 -> 打开**评教列表**界面

## 配置

1. 评价等级：`@value='A_优'` 
2. 评语内容：`COMMENT_TEXT = "无"`
3. 保存前的等待时间 (秒)：`WAIT_BEFORE_SAVE = 6`

## 声明

**该脚本仅用于学习用途**

**该脚本仅用于学习用途**

**该脚本仅用于学习用途**
