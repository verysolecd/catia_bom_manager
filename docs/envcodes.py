# 创建虚拟环境
python -m venv .venv
# 激活虚拟环境（Windows）
Set-ExecutionPolicy Bypass -Scope Process
.\.venv\Scripts\activate
.\.venv\Scripts\Activate.ps1
# 安装包，-e表示便携安装
pip install -e .

pyinstaller - -onefile - -add-data "resources/icons;resources/icons" src/main.py

pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple/

-i https: // pypi.tuna.tsinghua.edu.cn/simple/


目前国内比较好用的pypi源有：

http://mirrors.aliyun.com/pypi/simple/          阿里云

http://pypi.douban.com/simple/                     豆瓣

https://pypi.mirrors.ustc.edu.cn/simple/         中国科技大学

http://pypi.mirrors.opencas.cn/simple/           中科院

https://pypi.tuna.tsinghua.edu.cn/simple/       清华大学


GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG
# 初始化本地Git仓库
git init
# 添加远程仓库地址
git remote add origin <GitHub仓库地址>
# 拉取远程仓库代码（如果有）
git pull origin main


# 添加更改到暂存区
git add .
# 提交更改
git commit -m "调试并更新代码"
# 推送到GitHub
git push origin main


GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG