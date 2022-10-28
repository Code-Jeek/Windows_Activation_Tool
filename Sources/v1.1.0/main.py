"""导入模块"""
# 内置模块
import threading
import pythoncom
import ctypes
import sys
import os
import re

# 外置模块
from ttkbootstrap.constants import *
from urllib.request import urlopen
from json import load
import ttkbootstrap as ttk
import wmi


"""定义常量"""
WRITER = "Jeek"
VERSION = "1.1.0"
ACTIVATION_INFO_TEXT = """Windows激活工具支持产品数字激活和kms激活"""
OTHER_MESSAGES_INFO_TEXT = """请输入你的相关激活信息"""
VIEW_LICENSES_INFO_TEXT = """Windows激活工具将查看你的电脑许可证信息"""
WINDOWS_KEYS = [
    "VK7JG-NPHTM-C97JM-9MPGT-3V66T",
    "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    "TPYNC-4J6KF-4B4GP-2HD89-7XMP6",
    "2BXNW-6CGWX-9BXPV-YJ996-GMT6T",
    "NRTT2-86GJM-T969G-8BCBH-BDWXG",
    "XC88X-9N9QX-CDRVP-4XV22-RVV26",
    "TNM78-FJKXR-P26YV-GP8MB-JK8XG",
    "TR8NX-K7KPD-YTRW3-XTHKX-KQBP6",
    "NPPR9-FWDCX-D2C8J-H872K-2YT43",
    "NYW94-47Q7H-7X9TT-W7TXD-JTYPM",
    "NJ4MX-VQQ7Q-FP3DB-VDGHX-7XM87",
    "MH37W-N47XK-V7XM9-C7227-GCQG9"
    ]
WINDOWS_KMS = [
    "kms.03k.org",
    "kms.90zm.xyz",
    "kms.cangshui.net",
    "kms.myftp.org",
    "zh.us.to",
    "kms.chinancce.com",
    "kms.digiboy.ir",
    "kms.luody.info",
    "kms.mrxn.net",
    "kms8.MSGuides.com",
    "xykz.f3322.org",
    "kms.bige0.com",
    "kms.shuax.com",
    "kms9.MSGuides.com",
    "kms.lotro.cc",
    "www.ddddg.cn",
    "cy2617.jios.org",
    "enter.picp.net"
    ]

"""定义变量"""
menu_button_data = {
    1 : "self.form_activated_button",
    2 : "self.from_activated_button",
    4 : "self.form_view_licenses_button"
    }
window_data = {
    1 : [
        "self.activation_title",
        "self.activation_info",
        "self.Windows_keys_info",
        "self.Windows_keys_combobox",
        "self.Windows_kms_info",
        "self.Windows_kms_combobox",
        "self.other_messages_link",
        "self.activation_ok_button"
         ],
    2 : [
        "self.other_messages_reply_button",
        "self.other_messages_info",
        "self.type_keys_info",
        "self.type_keys_error",
        "self.type_keys_entry",
        "self.type_kms_info",
        "self.type_kms_error",
        "self.type_kms_entry"
        ],
    4 : [
        "self.view_licenses_title",
        "self.view_licenses_info",
        "self.computer_messages_text",
        "self.detailed_messages_link",
        "self.view_licenses_button",
        "self.fresh_computer_messages_button"
        ]
    }


class MainWindow():
    """用于GUI编写"""
    def __init__(self):
        """用于初始化主窗口"""
        self.window_state = None  # 初始化

        # 定义窗口
        self.root = ttk.Window(
            title="Windows激活工具",
            iconphoto=r".\resources\images\icon\icon.png",
            size=(900, 600),
            resizable=(False, False),
            themename="darkly",
            alpha=0.988
            )

        # 创建主菜单窗口组件
        self.menu_background = ttk.Frame(bootstyle=DARK, width=300, height=600)
        self.image = ttk.PhotoImage(file=r".\resources\images\Windows_logo.png")
        self.logo = ttk.Label(image=self.image, bootstyle=(DARK, INVERSE))
        self.writer = ttk.Label(text=f"作者: {WRITER}", bootstyle=(DARK, INVERSE))
        self.version = ttk.Label(text=f"版本: {VERSION}", bootstyle=(DARK, INVERSE))
        self.form_activated_button = ttk.Button(
            text="激活Windows许可证",
            bootstyle=(INFO, TOOLBUTTON),
            width=26,
            command=self.activation_arrang
            )
        self.form_view_licenses_button = ttk.Button(
            text="查看Windows许可证信息",
            bootstyle=(INFO, TOOLBUTTON),
            width=26,
            command=self.view_licenses_arrang
            )

        # 创建'激活Windows许可证'功能窗口组件
        self.activation_title = ttk.Label(text="激活Windows许可证", font=("微软雅黑", 22))
        self.activation_info = ttk.Label(text=ACTIVATION_INFO_TEXT, bootstyle=LIGHT, font=("微软雅黑", 12))
        self.Windows_keys_info = ttk.Label(text="请选择你的Windows产品密钥: ", font="微软雅黑")
        self.Windows_keys_combobox = ttk.Combobox(state=READONLY, values=WINDOWS_KEYS, width=41)
        self.Windows_kms_info = ttk.Label(text="请选择你的kms: ", font="微软雅黑")
        self.Windows_kms_combobox = ttk.Combobox(state=READONLY, values=WINDOWS_KMS, width=41)
        self.other_messages_link = ttk.Button(text="找不到我的相关信息?", bootstyle=(INFO, LINK), command=self.other_messages_arrang)
        self.activation_ok_button = ttk.Button(text="确  定", bootstyle=SUCCESS, width=8, command=self.activation)

        #创建'其它信息填写'功能窗口组件
        self.reply_image = ttk.PhotoImage(file=r".\resources\images\reply_light.png")
        self.other_messages_reply_button = ttk.Button(image=self.reply_image, bootstyle=(DARK, OUTLINE, TOOLBUTTON), command=self.recover_activation_arrang)
        self.other_messages_info = ttk.Label(text=OTHER_MESSAGES_INFO_TEXT, bootstyle=LIGHT, font=("微软雅黑", 12))
        self.type_keys_info = ttk.Label(text="请输入Windows产品密钥: ", font="微软雅黑")
        self.type_keys_error = ttk.Label(text="请输入正确的Windows产品密钥!", bootstyle=DANGER, font="微软雅黑")
        self.key_message = ttk.StringVar()
        self.type_keys_entry = ttk.Entry(textvariable=self.key_message, width=47)
        self.type_kms_info = ttk.Label(text="请输入Windows产品密钥: ", font="微软雅黑")
        self.type_kms_error = ttk.Label(text="请输入正确的kms服务器!", bootstyle=DANGER, font="微软雅黑")
        self.kms_message = ttk.StringVar()
        self.type_kms_entry = ttk.Entry(textvariable=self.kms_message, width=47)

        # 创建'查看Windows许可证信息'功能窗口组件
        self.view_licenses_title = ttk.Label(text="查看Windows许可证信息", font=("微软雅黑", 22))
        self.view_licenses_info = ttk.Label(text=VIEW_LICENSES_INFO_TEXT, bootstyle=LIGHT, font=("微软雅黑", 12))
        self.computer_messages_text = ttk.ScrolledText(state=DISABLED, wrap=WORD, font="微软雅黑", width=59, height=13)
        self.detailed_messages_link = ttk.Button(text="获取更详细的电脑信息", bootstyle=(INFO, LINK), command=self.detailed_messages)
        self.fresh_computer_messages_button = ttk.Button(text="刷  新", bootstyle=(PRIMARY, OUTLINE), width=8, command=self.fresh_computer_messages)
        self.view_licenses_button = ttk.Button(text="查  看", bootstyle=SUCCESS, width=8, command=self.view_licenses)

        self.insert_computer_messages()  # 插入电脑信息

    def main_menu_arrang(self):
        """用于画入主菜单窗口组件"""
        self.menu_background.grid(rowspan=15)
        self.logo.grid(column=0, row=0)
        self.writer.grid(column=0, row=14, sticky=SW)
        self.version.grid(column=0, row=14, sticky=SE)
        self.form_activated_button.grid(row=2, rowspan=15, ipady=16, sticky=N)
        self.form_view_licenses_button.grid(row=4, rowspan=15, ipady=16, sticky=N)
        
        self.root.mainloop()  # 显示窗口

    def activation_arrang(self):
        """用于画入'激活Windows许可证'功能窗口组件"""
        self.form_activated_button.config(state=DISABLED)  # 设置'激活Windows许可证'按钮状态为不显示
        self.recover_menu_button()  # 恢复主菜单窗口按钮状态

        self.update_window(1)  # 更新窗口

        self.activation_title.grid(column=1, columnspan=2, row=0, padx=18, sticky=W)
        self.activation_info.grid(column=1, columnspan=2, row=1, rowspan=15, padx=18, sticky=NW)
        self.Windows_keys_info.grid(column=1, row=3, rowspan=15, padx=18, sticky=NW)
        self.Windows_keys_combobox.grid(column=2, row=3, rowspan=15, padx=4, sticky=NW)
        self.Windows_kms_info.grid(column=1, row=5, rowspan=15, padx=18, sticky=NW)
        self.Windows_kms_combobox.grid(column=2, row=5, rowspan=15, padx=4, sticky=NW)
        self.other_messages_link.grid(column=1, row=6, rowspan=15, padx=18, sticky=W)
        self.activation_ok_button.grid(column=2, row=6, rowspan=15, padx=4, sticky=E)

        # 设置下拉框默认值
        self.Windows_keys_combobox.set(WINDOWS_KEYS[0])
        self.Windows_kms_combobox.set(WINDOWS_KMS[0])

    def other_messages_arrang(self):
        """用于画入'其它信息填写'功能窗口组件"""
        self.update_window(2)  # 更新窗口

        self.other_messages_reply_button.grid(column=1, columnspan=2, row=0, padx=18, sticky=W)
        self.other_messages_info.grid(column=1, columnspan=2, row=1, rowspan=15, padx=18, sticky=NW)
        self.type_keys_info.grid(column=1, row=3, rowspan=15, padx=18, sticky=NW)
        self.type_keys_entry.grid(column=2, row=3, rowspan=15, padx=4, sticky=NW)
        self.type_kms_info.grid(column=1, row=5, rowspan=15, padx=18, sticky=NW)
        self.type_kms_entry.grid(column=2, row=5, rowspan=15, padx=4, sticky=NW)
        self.activation_ok_button.grid(column=2, row=6, rowspan=15, padx=4, sticky=E)

    def view_licenses_arrang(self):
        """用于画入'查看Windows许可证信息'功能窗口组件"""
        self.form_view_licenses_button.config(state=DISABLED)  # 设置'查看Windows许可证信息'按钮状态禁用
        self.recover_menu_button()  # 恢复主菜单窗口按钮状态

        self.update_window(4)  # 更新窗口

        self.view_licenses_title.grid(column=1, columnspan=5, row=0, padx=18, sticky=W)
        self.view_licenses_info.grid(column=1, columnspan=5, row=1, rowspan=15, padx=18, sticky=NW)
        self.computer_messages_text.grid(column=1, columnspan=5, row=2, rowspan=15, padx=18, sticky=NW)
        self.detailed_messages_link.grid(column=1, columnspan=5, row=8, rowspan=15, padx=18, sticky=W)
        self.fresh_computer_messages_button.grid(column=5, row=8, rowspan=15, padx=18, sticky=W)
        self.view_licenses_button.grid(column=5, row=8, rowspan=15, padx=18, sticky=E)

    def update_window(self, data):
        """用于更新窗口"""
        if self.window_state is not None and self.window_state != data:
            for item in window_data[self.window_state]:
                eval(item).grid_forget()  # 隐藏对应功能窗口所有组件

        self.window_state = data  # 更改设置值

    def recover_menu_button(self):
        """用于恢复主菜单窗口按钮状态"""
        if self.window_state is not None:
            eval(menu_button_data[self.window_state]).config(state=NORMAL)  # 设置按钮状态为正常

    def recover_activation_arrang(self):
        """用于恢复'激活Windows许可证'功能窗口组件"""
        self.update_window(1)  # 更新窗口

        self.activation_title.grid(column=1, columnspan=2, row=0, padx=18, sticky=W)
        self.activation_info.grid(column=1, columnspan=2, row=1, rowspan=15, padx=18, sticky=NW)
        self.Windows_keys_info.grid(column=1, row=3, rowspan=15, padx=18, sticky=NW)
        self.Windows_keys_combobox.grid(column=2, row=3, rowspan=15, padx=4, sticky=NW)
        self.Windows_kms_info.grid(column=1, row=5, rowspan=15, padx=18, sticky=NW)
        self.Windows_kms_combobox.grid(column=2, row=5, rowspan=15, padx=4, sticky=NW)
        self.other_messages_link.grid(column=1, row=6, rowspan=15, padx=18, sticky=W)
        self.activation_ok_button.grid(column=2, row=6, rowspan=15, padx=4, sticky=E)


    def get_computer_messages(self):
        """用于获取电脑信息"""
        self.computer_messages_file = open("computer_messages.txt", "w+")  # 以读写模式打开文件

        pythoncom.CoInitialize()  # 初始化, 防止线程错误
        self.w = wmi.WMI()  # 初始化
        self.computer_messages_list = []  # 设定存放信息列表

        # 获取CPU信息
        data = []  # 设定临时存放列表
        for obj in self.w.Win32_Processor():  # 获取所有CPU信息
            self.computer_messages_file.write(str(obj))  # 将CPU信息写入文件
            data.append(f"{obj.Name}")  # 将CPU信息存入临时列表
        data[0] = "中央处理器(CPU): " + data[0]  # 防止类型转换错乱
        self.computer_messages_list.append(data)  # 将临时信息存入信息列表

        # 获取显卡信息
        data = []
        for obj in self.w.Win32_VideoController():
            self.computer_messages_file.write(str(obj))
            data.append(f"{obj.Name}")
        data[0] = "显卡(GPU): " + data[0]
        self.computer_messages_list.append(data)

        # 获取内存信息
        data = []
        size = 0
        for obj in self.w.Win32_PhysicalMemory():
            self.computer_messages_file.write(str(obj))
            size += int(obj.Capacity) // 1024 // 1024 // 1024
            data.append(f"{obj.Manufacturer} {obj.Caption} {obj.Speed}MHz {int(obj.Capacity) // 1024 // 1024 // 1024}GB")
        data.append(f"共{size}GB")
        data[0] = "内存: " + data[0]
        self.computer_messages_list.append(data)

        # 获取磁盘信息
        data = []
        size = 0
        for obj in self.w.Win32_DiskDrive():
            self.computer_messages_file.write(str(obj))
            size += int(obj.Size) // 1024 // 1024 // 1024
            data.append(f"{obj.Caption} {obj.BytesPerSector}GB {int(obj.Size) // 1024 // 1024 // 1024}GB可用")
        data.append(f"共{size}GB可用")
        data[0] = "磁盘: " + data[0]
        self.computer_messages_list.append(data)

        # 获取网卡信息
        data = []
        for obj in self.w.Win32_NetworkAdapterConfiguration(IPEnabled=True):
            self.computer_messages_file.write(str(obj))
            data.append(f"{obj.Description} 本地IPv4: {obj.IPAddress[0]} IPv6: {obj.IPAddress[1]}")
        data.append("本机公网IP: " + load(urlopen('http://jsonip.com'))['ip'])
        data[0] = "网卡: " + data[0]
        self.computer_messages_list.append(data)

        # 保存并关闭文件
        self.computer_messages_file.close()

        # 类型转换
        self.computer_messages_data = []  # 创建类型转换临时列表
        for item in self.computer_messages_list:
            self.computer_messages_data.append("\n".join(item))  # 将各电脑信息转换后存入临时列表

        self.computer_messages_str = "\n\n".join(self.computer_messages_data)  # 将临时列表转换为字符串

    def detailed_messages(self):
        """用于打开详细电脑信息文件"""
        def open_file():
            os.popen(f"start computer_messages.txt").read()  # 调用终端打开文件

            # 恢复按钮
            self.detailed_messages_link.config(state=NORMAL)
            self.fresh_computer_messages_button.config(state=NORMAL)

        # 设定按钮状态为禁用
        self.detailed_messages_link.config(state=DISABLED)
        self.fresh_computer_messages_button.config(state=DISABLED)

        file_thread = threading.Thread(target=open_file)
        file_thread.start()  # 开始进程

    def insert_computer_messages(self):
        """用于插入电脑信息"""
        def insert_threading():
            # 设定相关按钮状态为禁用
            self.detailed_messages_link.config(state=DISABLED)
            self.fresh_computer_messages_button.config(state=DISABLED)

            # 写入加载内容
            self.computer_messages_text.config(state=NORMAL)  # 设置允许写入
            self.computer_messages_text.delete("0.0", END)  # 清空内容
            self.computer_messages_text.insert(INSERT, "获取信息中...")  # 写入加载文本
            self.computer_messages_text.config(state=DISABLED)  # 禁止写入
            self.get_computer_messages()  # 获取电脑信息

            self.computer_messages_text.config(state=NORMAL)  # 设置允许写入
            self.computer_messages_text.delete("0.0", END)  # 清空内容
            self.computer_messages_text.insert(INSERT, self.computer_messages_str)  # 写入电脑信息
            self.computer_messages_text.config(state=DISABLED)  # 禁止写入

            # 恢复相关按钮
            self.detailed_messages_link.config(state=NORMAL)
            self.fresh_computer_messages_button.config(state=NORMAL)

        # 创建获取进程
        insert_computer_messages_thread = threading.Thread(target=insert_threading)
        insert_computer_messages_thread.start()  # 开始进程

    def fresh_computer_messages(self):
        """用于刷新电脑信息"""
        self.insert_computer_messages()

    def activation(self):
        """用于激活Windows许可证"""
        def activing(key=None, kms=None):
            """用于调用终端激活Windows许可证"""
            if key != None:
                os.popen(f"slmgr /ipk {key}").read()  # 设定产品密钥
            if kms != None:
                os.popen(f"slmgr /skms {kms}").read()  # 设定kms
            os.popen("slmgr /ato").read()  # 激活Windows许可证

            # 恢复按钮
            self.activation_ok_button.config(state=NORMAL)

        # 设置按钮状态为禁用
        self.activation_ok_button.config(state=DISABLED)

        # 获取相关信息
        if self.window_state == 1:
            key = self.Windows_keys_combobox.get()
            kms = self.Windows_kms_combobox.get()
        else:
            key = self.key_message.get()
            kms = self.kms_message.get()

        self.match_key = re.match(r"^([A-Za-z0-9]{5}-){4}[A-Za-z0-9]{5}$", key)
        self.match_kms = re.match(r"^([A-Za-z0-9]+.)+[A-Za-z0-9]$", kms)

        # 检查相关信息
        if self.match_key == None:
            self.type_keys_error.grid(column=2, row=4, rowspan=15, sticky=NW)  # 画入错误提示
            # 恢复按钮
            self.activation_ok_button.config(state=NORMAL)

            if self.match_kms == None:
                self.type_kms_error.grid(column=2, row=6, rowspan=15, sticky=NW)  # 画入错误提示
                # 恢复按钮
                self.activation_ok_button.config(state=NORMAL)
                return
            else:
                self.type_kms_error.grid_forget()  # 删除错误提示
        else:
            self.type_keys_error.grid_forget()  # 删除错误提示

        if self.match_kms == None:
            self.type_kms_error.grid(column=2, row=6, rowspan=15, sticky=NW)  # 画入错误提示
            # 恢复按钮
            self.activation_ok_button.config(state=NORMAL)
            return
        else:
            self.type_kms_error.grid_forget()  # 删除错误提示

        # 创建进程
        self.activing_thread = threading.Thread(target=activing, args=(key, kms))
        self.activing_thread.start()  # 开始进程

    def view_licenses(self):
        """用于查看Windows许可证信息"""
        def viewing_licenses():
            os.popen("slmgr /dlv").read()

            # 恢复按钮
            self.view_licenses_button.config(state=NORMAL)

        # 设置按钮状态为禁用
        self.view_licenses_button.config(state=DISABLED)

        # 创建进程
        view_licenses_thread = threading.Thread(target=viewing_licenses)
        view_licenses_thread.start()  # 开始进程


def is_admin():
    """用于获取管理员权限"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    """主函数"""
    os.chdir(os.path.dirname(__file__))  # 将程序运行目录修改为文件所在目录

    root = MainWindow()
    root.main_menu_arrang()

if __name__ == "__main__":
    if is_admin():
        main()
    else:
        if sys.version_info[0] == 3:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
