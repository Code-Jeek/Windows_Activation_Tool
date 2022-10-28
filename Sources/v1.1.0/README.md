# Windows_Activation_Tool-v1.1.0

## 环境依赖

- Python3.x
- ttkbootstrap
- wmi

## 部署步骤

1. ### 安装Python3.x

[Python官网](https://www.python.org)下载安装

2. ### 安装ttkbootstrap库

```
python -m pip install ttkbootstrap
python -m pip install git+https://github.com/israel-dryer/ttkbootstrap
```

3. ### 安装wmi库

`python -m pip install wmi`

4. ### 测试环境

在Shell中输入`Python`成功进入Python环境说明Python安装成功; 
在Python环境中导入ttkbootstrap和wmi库, 不报错说明外置库安装成功

## 目录结构描述

```
Windows_Activation_Tool:
│  computer_messages.txt
│  main.py
│  README.md
│  
└─resources
    └─images
        │  reply_dark.png
        │  reply_light.png
        │  Windows_logo.png
        │  
        └─icon
                icon.ico
                icon.png
```

## [v1.1.0] - 2022-10-28

### 新增

- 在`激活Windows许可证`功能窗口新增了`找不到我的相关信息?`按钮引导的`其它信息填写`功能窗口
- 将主窗口设为半透明(透明度很小)

### 优化

- 优化了目录结构

###  删除

- 删除了屏幕居中显示
