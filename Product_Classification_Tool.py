import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime

def cn_number_to_arabic(cn_number):
    cn_nums = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
    cn_units = {'十': 10, '百': 100, '千': 1000, '万': 10000}
    
    if not cn_number:
        return 0
    
    result = 0
    unit = 1
    for char in reversed(cn_number):
        if char in cn_units:
            unit = cn_units[char]
        elif char in cn_nums:
            result += cn_nums[char] * unit
            unit = 1
    
    return result

def validate_date(date_str):
    try:
        if isinstance(date_str, str):
            # 尝试解析日期字符串
            pd.to_datetime(date_str)
            return True
    except:
        pass
    return False

class ProductClassificationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("产品分类工具")
        self.root.geometry("800x600")

        # 创建主框架
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 创建文件选择区域
        self.create_file_selection_area()

        # 创建日志区域
        self.create_log_area()

        # 初始化文件路径变量
        self.file_paths = []
        self.save_path = None

    def create_file_selection_area(self):
        # 文件选择框架
        file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding=10)
        file_frame.pack(fill=tk.X, pady=5)

        # 选择文件按钮
        ttk.Button(file_frame, text="选择文件", command=self.select_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="选择保存位置", command=self.select_save_location).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="开始处理", command=self.process_files).pack(side=tk.LEFT, padx=5)

    def create_log_area(self):
        # 日志框架
        log_frame = ttk.LabelFrame(self.main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 创建文本框和滚动条
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=20)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # 放置文本框和滚动条
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if files:
            self.file_paths = files
            self.log(f"已选择 {len(files)} 个文件")
            for file in files:
                self.log(f"- {os.path.basename(file)}")

    def select_save_location(self):
        save_path = filedialog.askdirectory(title="选择保存位置")
        if save_path:
            self.save_path = save_path
            self.log(f"保存位置: {save_path}")

    def process_files(self):
        if not self.file_paths:
            messagebox.showerror("错误", "请先选择要处理的文件！")
            return

        if not self.save_path:
            messagebox.showerror("错误", "请选择保存位置！")
            return

        try:
            for file_path in self.file_paths:
                self.process_single_file(file_path)
            
            messagebox.showinfo("成功", "所有文件处理完成！")
        except Exception as e:
            messagebox.showerror("错误", f"处理文件时出错：{str(e)}")

    def process_single_file(self, file_path):
        self.log(f"正在处理文件: {os.path.basename(file_path)}")

        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查必要的列是否存在
        required_columns = ['产品名称', '规格']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"文件中缺少必要的列：{col}")

        # 添加新列
        df['产品分类'] = ''
        df['规格数值'] = ''
        df['单位'] = ''

        # 处理每一行
        for index, row in df.iterrows():
            product_name = str(row['产品名称'])
            spec = str(row['规格'])

            # 提取规格数值和单位
            spec_value = ''
            unit = ''

            # 处理规格信息
            if 'kg' in spec.lower() or '公斤' in spec:
                unit = 'kg'
                spec_value = spec.lower().replace('kg', '').replace('公斤', '').strip()
            elif 'g' in spec.lower() or '克' in spec:
                unit = 'g'
                spec_value = spec.lower().replace('g', '').replace('克', '').strip()
            elif '包' in spec:
                unit = '包'
                spec_value = spec.replace('包', '').strip()

            # 转换中文数字
            if spec_value:
                try:
                    if any(char in spec_value for char in ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']):
                        spec_value = str(cn_number_to_arabic(spec_value))
                except:
                    pass

            # 更新DataFrame
            df.at[index, '规格数值'] = spec_value
            df.at[index, '单位'] = unit

            # 确定产品分类
            if '鱼' in product_name:
                df.at[index, '产品分类'] = '海鲜'
            elif '虾' in product_name:
                df.at[index, '产品分类'] = '海鲜'
            elif '蟹' in product_name:
                df.at[index, '产品分类'] = '海鲜'
            elif '贝' in product_name:
                df.at[index, '产品分类'] = '海鲜'
            elif '肉' in product_name:
                df.at[index, '产品分类'] = '肉类'
            elif '菜' in product_name:
                df.at[index, '产品分类'] = '蔬菜'
            elif '豆' in product_name:
                df.at[index, '产品分类'] = '豆制品'
            else:
                df.at[index, '产品分类'] = '其他'

        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"产品分类_{timestamp}.xlsx"
        output_path = os.path.join(self.save_path, output_filename)

        # 保存处理后的文件
        df.to_excel(output_path, index=False)
        self.log(f"文件处理完成，已保存至：{output_path}")
