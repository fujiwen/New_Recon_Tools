import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import pandas as pd
from Product_Classification_Tool import ProductClassificationApp

class BldBuyApp:
    def __init__(self):
        self.root = ttk.Window(title="对账工具", themename="sandstone")
        self.root.geometry("800x653")
        self.root.resizable(False, False)

        # 设置窗口图标
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "favicon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        # 确保配置文件存在
        self.config_file = "config.txt"
        self.ensure_config_exists()

        # 加载主题
        self.theme = self.load_theme()
        self.style = ttk.Style()
        self.style.theme_use(self.theme)

        # 创建主题选择器
        self.create_theme_selector()

        # 创建左侧功能按钮
        self.create_left_buttons()

        # 创建日志区域
        self.create_log_area()

        # 初始化文件路径变量
        self.file_path = None
        self.save_path = None

        # 检查过期时间
        self.check_expiry()

    def check_expiry(self):
        expiry_date = datetime(2024, 12, 31)
        current_date = datetime.now()
        if current_date > expiry_date:
            messagebox.showerror("错误", "软件已过期，请联系管理员更新！")
            self.root.destroy()
            sys.exit()

    def ensure_config_exists(self):
        if not os.path.exists(self.config_file):
            default_config = """B2:海口索菲特大酒店

D2:海南省海口市龙华区滨海大道105号

E2:符小瑜 0898-31289999

B32:abbyfu@hksft.com

hotelname:海口索菲特大酒店

Sheet_tittle:供货明细表

theme:sandstone"""
            with open(self.config_file, "w", encoding="utf-8") as f:
                f.write(default_config)

    def load_theme(self):
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                config = f.read()
                for line in config.split("\n"):
                    if line.startswith("theme:"):
                        return line.split(":")[1].strip()
        except Exception as e:
            print(f"加载主题时出错: {e}")
        return "sandstone"

    def save_theme(self, theme):
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                lines = f.readlines()
            
            with open(self.config_file, "w", encoding="utf-8") as f:
                theme_found = False
                for line in lines:
                    if line.startswith("theme:"):
                        f.write(f"theme:{theme}\n")
                        theme_found = True
                    else:
                        f.write(line)
                if not theme_found:
                    f.write(f"\ntheme:{theme}")
        except Exception as e:
            print(f"保存主题时出错: {e}")

    def create_theme_selector(self):
        # 创建主题选择下拉框
        themes_frame = ttk.Frame(self.root)
        themes_frame.pack(side="top", fill="x", padx=5, pady=5)

        ttk.Label(themes_frame, text="主题：").pack(side="left")
        self.theme_var = tk.StringVar(value=self.theme)
        theme_combo = ttk.Combobox(themes_frame, textvariable=self.theme_var, values=self.style.theme_names())
        theme_combo.pack(side="left")
        theme_combo.bind("<<ComboboxSelected>>", self.change_theme)

    def change_theme(self, event):
        selected_theme = self.theme_var.get()
        self.style.theme_use(selected_theme)
        self.save_theme(selected_theme)

    def create_left_buttons(self):
        # 创建左侧按钮框架
        left_frame = ttk.Frame(self.root)
        left_frame.pack(side="left", fill="y", padx=5, pady=5)

        # 创建按钮
        ttk.Button(left_frame, text="对账工具", command=self.open_recon_tool, width=15).pack(pady=5)
        ttk.Button(left_frame, text="产品分类工具", command=self.open_classification_tool, width=15).pack(pady=5)

    def create_log_area(self):
        # 创建日志框架
        log_frame = ttk.Frame(self.root)
        log_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        log_frame.configure(height=300)

        # 创建日志文本框和滚动条
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=20)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # 放置日志文本框和滚动条
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

    def open_recon_tool(self):
        self.log("正在打开对账工具...")
        # 在这里添加对账工具的具体实现
        self.log("对账工具已打开")

    def open_classification_tool(self):
        self.log("正在打开产品分类工具...")
        classification_window = tk.Toplevel(self.root)
        app = ProductClassificationApp(classification_window)
        self.log("产品分类工具已打开")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = BldBuyApp()
    app.run()