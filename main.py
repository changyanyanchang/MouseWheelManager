import sys
import time  # 新增：用于重启设备时的短暂延时
import ctypes
from ctypes import wintypes
import winreg
import customtkinter as ctk
from tkinter import messagebox

# =========================================================================
# 1. 底层 CfgMgr32 定义
# =========================================================================

cfgmgr32 = ctypes.windll.cfgmgr32

CR_SUCCESS = 0x00000000
DN_STARTED = 0x00000008
# 鼠标设备的专属 GUID
MOUSE_CLASS_GUID = "{4D36E96F-E325-11CE-BFC1-08002BE10318}"

class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong), ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort), ("Data4", ctypes.c_ubyte * 8)]

class DEVPROPKEY(ctypes.Structure):
    _fields_ = [("fmtid", GUID), ("pid", ctypes.c_ulong)]

# DEVPKEY_Device_FriendlyName
guid_friendly = GUID(0xa45c254e, 0xdf1c, 0x4efd, (ctypes.c_ubyte * 8)(0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0))
DEVPKEY_FriendlyName = DEVPROPKEY(guid_friendly, 14)

# === 新增：定义禁用/启用设备的 API ===
try:
    # 定义函数原型，防止参数传递错误
    cfgmgr32.CM_Disable_DevNode.argtypes = [ctypes.c_ulong, ctypes.c_ulong]
    cfgmgr32.CM_Disable_DevNode.restype = ctypes.c_ulong

    cfgmgr32.CM_Enable_DevNode.argtypes = [ctypes.c_ulong, ctypes.c_ulong]
    cfgmgr32.CM_Enable_DevNode.restype = ctypes.c_ulong
except AttributeError:
    pass

# =========================================================================
# 2. 注册表与设备助手类
# =========================================================================

class RegistryHelper:
    # === CfgMgr32 辅助函数 START ===
    @staticmethod
    def get_devnode_status(pnp_id: str):
        """检查设备是否连接，并返回 dev_inst 句柄"""
        dev_inst = ctypes.c_ulong()
        ret = cfgmgr32.CM_Locate_DevNodeW(ctypes.byref(dev_inst), ctypes.c_wchar_p(pnp_id), 0)
        if ret != CR_SUCCESS:
            return False, 0
        
        status = ctypes.c_ulong()
        problem = ctypes.c_ulong()
        ret = cfgmgr32.CM_Get_DevNode_Status(ctypes.byref(status), ctypes.byref(problem), dev_inst, 0)
        return bool(status.value & DN_STARTED), dev_inst.value

    @staticmethod
    def restart_device(pnp_id: str) -> bool:
        """
        【新增核心功能】软件重启设备：禁用 -> 等待 -> 启用。
        这会强制驱动重新读取注册表配置，无需物理插拔。
        """
        dev_inst = ctypes.c_ulong()
        # 1. 重新获取句柄 (确保句柄是最新的)
        ret = cfgmgr32.CM_Locate_DevNodeW(ctypes.byref(dev_inst), ctypes.c_wchar_p(pnp_id), 0)
        if ret != CR_SUCCESS:
            print(f"Error: 无法定位设备 {pnp_id}")
            return False

        # 2. 禁用设备 (相当于在设备管理器右键禁用)
        ret = cfgmgr32.CM_Disable_DevNode(dev_inst, 0)
        if ret != CR_SUCCESS:
            print(f"Error: 禁用设备失败，错误代码: {ret}")
            return False
        
        # 给系统一点喘息时间，防止状态切换过快导致死锁或未生效
        time.sleep(1.0) 

        # 3. 启用设备 (驱动重新初始化，读取注册表)
        ret = cfgmgr32.CM_Enable_DevNode(dev_inst, 0)
        if ret != CR_SUCCESS:
            print(f"Error: 启用设备失败，错误代码: {ret}")
            return False

        return True

    @staticmethod
    def get_property(dev_inst: int, property_key: DEVPROPKEY) -> str:
        prop_type = ctypes.c_ulong()
        buffer_size = ctypes.c_ulong(0)
        cfgmgr32.CM_Get_DevNode_PropertyW(dev_inst, ctypes.byref(property_key), ctypes.byref(prop_type), None, ctypes.byref(buffer_size), 0)
        if buffer_size.value == 0: return ""
        buffer = ctypes.create_unicode_buffer(buffer_size.value)
        ret = cfgmgr32.CM_Get_DevNode_PropertyW(dev_inst, ctypes.byref(property_key), ctypes.byref(prop_type), buffer, ctypes.byref(buffer_size), 0)
        return buffer.value if ret == CR_SUCCESS else ""

    @staticmethod
    def get_parent_handle(child_inst: int) -> int:
        parent_inst = ctypes.c_ulong()
        ret = cfgmgr32.CM_Get_Parent(ctypes.byref(parent_inst), child_inst, 0)
        return parent_inst.value if ret == CR_SUCCESS else 0

    @staticmethod
    def get_device_id_from_handle(dev_inst: int) -> str:
        ptr_buf = ctypes.create_unicode_buffer(200)
        ret = cfgmgr32.CM_Get_Device_IDW(dev_inst, ptr_buf, 200, 0)
        return ptr_buf.value if ret == CR_SUCCESS else ""

    @staticmethod
    def find_real_name_via_parent(dev_inst: int, current_pnp_id: str, default_desc: str) -> str:
        """核心逻辑：向上查找父节点以获取真实硬件名称"""
        # 1. 尝试直接获取 FriendlyName
        friendly = RegistryHelper.get_property(dev_inst, DEVPKEY_FriendlyName)
        
        # 如果是 HID/BTH 设备，尝试往上找"爸爸"
        if "HID" in current_pnp_id.upper() or "BTH" in current_pnp_id.upper():
            curr_inst = dev_inst
            for _ in range(3): # 最多往上找3层
                parent_inst = RegistryHelper.get_parent_handle(curr_inst)
                if parent_inst == 0: break
                
                parent_friendly = RegistryHelper.get_property(parent_inst, DEVPKEY_FriendlyName)
                
                # 如果父节点有友好名称，且不是通用的"枚举器"
                if parent_friendly and ("ENUMERATOR" not in parent_friendly.upper()):
                     return parent_friendly
                
                curr_inst = parent_inst

        if friendly:
            return friendly
        return default_desc
    # === CfgMgr32 辅助函数 END ===

    @staticmethod
    def get_registry_value_safe(key, value_name):
        try:
            value, _ = winreg.QueryValueEx(key, value_name)
            return value
        except FileNotFoundError:
            return None

    @staticmethod
    def scan_mice():
        """扫描 HID 和 Bluetooth 总线下的鼠标设备"""
        devices = []
        # 扫描 HID 和 BTH (覆盖 USB 接收器和 纯蓝牙鼠标)
        bus_list = ["HID", "BTH", "BTHENUM"]
        seen_ids = set()

        for bus in bus_list:
            base_path = f"SYSTEM\\CurrentControlSet\\Enum\\{bus}"
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path) as bus_key:
                    i = 0
                    while True:
                        try:
                            device_id_key_name = winreg.EnumKey(bus_key, i)
                            i += 1
                            with winreg.OpenKey(bus_key, device_id_key_name) as device_id_key:
                                j = 0
                                while True:
                                    try:
                                        instance_name = winreg.EnumKey(device_id_key, j)
                                        j += 1
                                        
                                        full_pnp_id = f"{bus}\\{device_id_key_name}\\{instance_name}"
                                        
                                        # 1. 检查连接状态 (底层 API)
                                        is_connected, dev_inst = RegistryHelper.get_devnode_status(full_pnp_id)
                                        if not is_connected:
                                            continue
                                        
                                        # 2. 打开注册表检查 ClassGUID
                                        try:
                                            # 构建完整的注册表路径
                                            reg_path_full = f"{base_path}\\{device_id_key_name}\\{instance_name}"
                                            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path_full) as instance_key:
                                                class_guid = RegistryHelper.get_registry_value_safe(instance_key, "ClassGUID")
                                                
                                                # 严格过滤：只保留鼠标 ClassGUID
                                                if not class_guid or class_guid.upper() != MOUSE_CLASS_GUID:
                                                    continue

                                                # 3. 检查是否有 Device Parameters\FlipFlopWheel
                                                # 只有有这个参数的鼠标，我们才能修改滚轮方向
                                                param_path_rel = f"{reg_path_full}\\Device Parameters"
                                                try:
                                                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, param_path_rel) as param_key:
                                                        winreg.QueryValueEx(param_key, "FlipFlopWheel")
                                                except FileNotFoundError:
                                                    # 虽然是鼠标且在线，但没有滚轮反转参数，跳过
                                                    continue

                                                # 4. 获取基础名称
                                                reg_desc = RegistryHelper.get_registry_value_safe(instance_key, "DeviceDesc")
                                                reg_friendly = RegistryHelper.get_registry_value_safe(instance_key, "FriendlyName")
                                                base_name = reg_friendly if reg_friendly else reg_desc
                                                if base_name and ";" in base_name:
                                                    base_name = base_name.split(";")[-1]
                                                
                                                # 5. 【核心】向上查找获取真实名称 (如 "MX Master 3S")
                                                real_name = RegistryHelper.find_real_name_via_parent(dev_inst, full_pnp_id, base_name)

                                                # 6. 过滤虚拟设备
                                                if "Terminal Server" in real_name or "Remote Desktop" in real_name:
                                                    continue

                                                # 7. 准备 UI 数据
                                                # 简单的 ID 显示
                                                id_display = full_pnp_id.split("\\")[1] 
                                                if "VID" in id_display and "PID" in id_display:
                                                    # 尝试简化显示
                                                    parts = id_display.split("&")
                                                    vid = [p for p in parts if "VID" in p]
                                                    pid = [p for p in parts if "PID" in p]
                                                    if vid and pid:
                                                        id_display = f"{vid[0]} {pid[0]}"

                                                if full_pnp_id not in seen_ids:
                                                    seen_ids.add(full_pnp_id)
                                                    devices.append({
                                                        "name": real_name,
                                                        "id_display": id_display,
                                                        "reg_path": param_path_rel, # 存储参数路径供后续修改
                                                        "pnp_id": full_pnp_id
                                                    })

                                        except OSError:
                                            continue

                                    except OSError:
                                        break
                        except OSError:
                            break
            except FileNotFoundError:
                continue
        
        # 排序：名字长的（通常是具体型号）排前面，"HID-compliant mouse" 排后面
        devices.sort(key=lambda x: len(x['name']), reverse=True)
        return devices

    @staticmethod
    def get_state(reg_path):
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_READ)
            val, _ = winreg.QueryValueEx(key, "FlipFlopWheel")
            winreg.CloseKey(key)
            return val
        except: return 0

    @staticmethod
    def set_state(reg_path, value):
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "FlipFlopWheel", 0, winreg.REG_DWORD, value)
            winreg.CloseKey(key)
            return True
        except: return False

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def run_as_admin():
    executable = sys.executable.replace("python.exe", "pythonw.exe")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, " ".join(sys.argv), None, 0)

# =========================================================================
# UI 部分
# =========================================================================

# 定义设计系统颜色和字体
THEME = {
    "bg_left": ("#F3F4F6", "#1F2937"),      # 侧边栏背景
    "bg_right": ("#FFFFFF", "#111827"),     # 主区域背景
    "accent": "#3B82F6",                    # 强调色 (蓝)
    "accent_hover": "#2563EB",
    "text_main": ("#111827", "#F9FAFB"),    # 主文本
    "text_sub": ("#6B7280", "#9CA3AF"),     # 副文本
    "card_bg": ("#F9FAFB", "#374151"),      # 卡片背景
    "success": "#10B981",                   # 成功/Win模式
    "warning": "#F59E0B",                   # 警告/Mac模式
    "list_hover": ("#E5E7EB", "#374151"),   # 列表悬停
    "list_selected": ("#DBEAFE", "#1E3A8A") # 列表选中
}

FONT_MAIN = ("Segoe UI", 14)
FONT_BOLD = ("Segoe UI", 14, "bold")
FONT_TITLE = ("Segoe UI", 24, "bold")
FONT_SUB = ("Segoe UI", 12)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mouse Wheel Manager")
        self.geometry("800x550")
        
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.devices = []
        self.selected_device = None
        self.device_buttons = []

        self.setup_layout()
        self.refresh_list()

    def setup_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # === 左侧边栏 ===
        self.left_frame = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=THEME["bg_left"])
        self.left_frame.grid(row=0, column=0, sticky="nsew")
        self.left_frame.grid_rowconfigure(3, weight=1)

        self.brand_label = ctk.CTkLabel(
            self.left_frame, 
            text="🖱️ 鼠标配置", 
            font=("Segoe UI", 20, "bold"),
            text_color=THEME["text_main"]
        )
        self.brand_label.grid(row=0, column=0, padx=20, pady=(30, 20), sticky="w")

        self.btn_refresh = ctk.CTkButton(
            self.left_frame, 
            text="🔄  刷新列表", 
            font=FONT_BOLD,
            height=40,
            fg_color=THEME["accent"],
            hover_color=THEME["accent_hover"],
            corner_radius=8,
            command=self.refresh_list
        )
        self.btn_refresh.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")

        self.lbl_list_header = ctk.CTkLabel(
            self.left_frame, text="在线设备", font=FONT_SUB, text_color=THEME["text_sub"], anchor="w"
        )
        self.lbl_list_header.grid(row=2, column=0, padx=20, pady=(0,5), sticky="nw")

        self.scroll_frame = ctk.CTkScrollableFrame(
            self.left_frame, 
            label_text="", 
            fg_color="transparent"
        )
        self.scroll_frame.grid(row=3, column=0, padx=10, pady=(0, 20), sticky="nsew")

        # === 右侧主内容区 ===
        self.right_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=THEME["bg_right"])
        self.right_frame.grid(row=0, column=1, sticky="nsew")
        
        self.empty_state = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.empty_state.place(relx=0.5, rely=0.5, anchor="center")
        ctk.CTkLabel(self.empty_state, text="👈", font=("Segoe UI", 48)).pack()
        ctk.CTkLabel(self.empty_state, text="请在左侧选择一个设备\n以开始配置", font=FONT_MAIN, text_color=THEME["text_sub"]).pack(pady=10)

        self.content_area = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        
        self.info_frame = ctk.CTkFrame(self.content_area, fg_color=THEME["card_bg"], corner_radius=12)
        self.info_frame.pack(fill="x", pady=(40, 20), padx=40)
        
        self.lbl_name = ctk.CTkLabel(self.info_frame, text="Device Name", font=FONT_TITLE, text_color=THEME["text_main"], anchor="w")
        self.lbl_name.pack(padx=20, pady=(20, 5), fill="x")
        
        self.lbl_id = ctk.CTkLabel(self.info_frame, text="VID:PID", font=("Consolas", 12), text_color=THEME["text_sub"], anchor="w")
        self.lbl_id.pack(padx=20, pady=(0, 20), fill="x")

        self.status_container = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.status_container.pack(fill="x", padx=40, pady=10)
        
        ctk.CTkLabel(self.status_container, text="当前滚轮行为", font=FONT_BOLD, text_color=THEME["text_main"]).pack(anchor="w", pady=(0,10))
        
        self.status_indicator = ctk.CTkButton(
            self.status_container,
            text="--",
            font=("Segoe UI", 16, "bold"),
            height=60,
            corner_radius=10,
            fg_color=THEME["list_hover"],
            text_color_disabled=THEME["text_main"],
            state="disabled"
        )
        self.status_indicator.pack(fill="x")

        self.action_container = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.action_container.pack(fill="x", padx=40, pady=30)
        
        ctk.CTkLabel(self.action_container, text="修改设置", font=FONT_BOLD, text_color=THEME["text_main"]).pack(anchor="w", pady=(0,10))

        self.btn_mac = ctk.CTkButton(
            self.action_container, 
            text="🍎 Mac 模式\n(自然滚动/反转)", 
            font=FONT_MAIN,
            height=80,
            fg_color=THEME["bg_left"],
            border_width=2,
            border_color=THEME["bg_left"],
            text_color=THEME["text_main"],
            hover_color=THEME["list_hover"],
            corner_radius=10,
            command=lambda: self.apply_setting(1)
        )
        self.btn_mac.pack(fill="x", pady=5)

        self.btn_win = ctk.CTkButton(
            self.action_container, 
            text="🪟 Windows 模式\n(传统滚动/默认)", 
            font=FONT_MAIN,
            height=80,
            fg_color=THEME["bg_left"],
            border_width=2,
            border_color=THEME["bg_left"],
            text_color=THEME["text_main"],
            hover_color=THEME["list_hover"],
            corner_radius=10,
            command=lambda: self.apply_setting(0)
        )
        self.btn_win.pack(fill="x", pady=5)

        # 提示文案修改
        self.lbl_hint = ctk.CTkLabel(
            self.content_area, 
            text="ℹ️ 修改设置后，鼠标可能会暂停响应 1 秒以自动重启设备", 
            font=("Segoe UI", 11), 
            text_color=THEME["text_sub"]
        )
        self.lbl_hint.pack(side="bottom", pady=20)

    def refresh_list(self):
        for widget in self.scroll_frame.winfo_children(): widget.destroy()
        self.device_buttons.clear()
        
        self.devices = RegistryHelper.scan_mice()
        
        if not self.devices:
            ctk.CTkLabel(self.scroll_frame, text="未检测到在线设备\n请检查连接", text_color=THEME["text_sub"]).pack(pady=20)
            return

        for i, dev in enumerate(self.devices):
            is_apple = "Apple" in dev['name']
            icon = "🍎" if is_apple else "🖱️"
            
            display_name = dev['name']
            
            btn = ctk.CTkButton(
                self.scroll_frame,
                text=f"{icon}  {display_name}\n      {dev['id_display']}",
                font=("Segoe UI", 13),
                anchor="w",
                height=60,
                fg_color="transparent",
                text_color=THEME["text_main"],
                hover_color=THEME["list_hover"],
                corner_radius=6,
                command=lambda idx=i: self.select_device(idx)
            )
            btn.pack(fill="x", pady=2, padx=5)
            self.device_buttons.append(btn)

    def select_device(self, index):
        self.selected_device = self.devices[index]
        
        for i, btn in enumerate(self.device_buttons):
            if i == index:
                btn.configure(fg_color=THEME["list_selected"], text_color=THEME["accent"])
            else:
                btn.configure(fg_color="transparent", text_color=THEME["text_main"])

        self.empty_state.place_forget()
        self.content_area.pack(fill="both", expand=True)

        self.lbl_name.configure(text=self.selected_device['name'])
        self.lbl_id.configure(text=self.selected_device['pnp_id'])
        
        self.update_status_ui()

    def update_status_ui(self):
        val = RegistryHelper.get_state(self.selected_device['reg_path'])
        hide_border_color = THEME["bg_left"]

        if val == 1:
            # === 场景：Mac 模式激活 ===
            self.status_indicator.configure(
                text="已启用：Mac 自然滚动", 
                fg_color=THEME["warning"], 
                text_color="#FFFFFF"
            )
            self.btn_mac.configure(
                border_color=THEME["warning"], 
                fg_color=THEME["bg_left"], 
                text_color=THEME["warning"]
            )
            self.btn_win.configure(
                border_color=hide_border_color, 
                fg_color=THEME["bg_left"], 
                text_color=THEME["text_sub"]
            )
        else:
            # === 场景：Windows 模式激活 ===
            self.status_indicator.configure(
                text="已启用：Windows 默认滚动", 
                fg_color=THEME["success"], 
                text_color="#FFFFFF"
            )
            self.btn_mac.configure(
                border_color=hide_border_color, 
                fg_color=THEME["bg_left"], 
                text_color=THEME["text_sub"]
            )
            self.btn_win.configure(
                border_color=THEME["success"], 
                fg_color=THEME["bg_left"], 
                text_color=THEME["success"]
            )

    def apply_setting(self, val):
        pnp_id = self.selected_device['pnp_id']
        reg_path = self.selected_device['reg_path']
        
        # 1. 写入注册表
        if RegistryHelper.set_state(reg_path, val):
            self.update_status_ui()
            
            # 2. 尝试软件重启设备
            # 更改鼠标光标为“忙碌”状态，提示用户正在处理
            self.configure(cursor="watch") 
            self.update() # 强制刷新 UI 以显示忙碌光标
            
            success_restart = RegistryHelper.restart_device(pnp_id)
            
            self.configure(cursor="") # 恢复鼠标光标
            
            if success_restart:
                # 成功：无需插拔
                messagebox.showinfo(
                    "设置已生效", 
                    "✅ 配置已更新并自动重载！\n\n设备已在后台自动重启（Disable -> Enable），新设置已立即生效。"
                )
            else:
                # 失败（极少情况，例如权限不足或设备正忙）：回退到提示插拔
                messagebox.showwarning(
                    "设置已保存", 
                    "✅ 配置已写入注册表。\n\n⚠️ 自动重启设备失败，请手动拔插鼠标接收器以生效。"
                )
        else:
            messagebox.showerror("错误", "无法写入注册表，请确保以管理员权限运行。")

if __name__ == "__main__":
    if is_admin():
        app = App()
        app.mainloop()
    else:
        run_as_admin()