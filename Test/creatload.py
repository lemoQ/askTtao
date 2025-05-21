import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import os


def generate_inps(excel_path, template_paths, output_folder):
    try:
        # 创建基于Excel文件名的子目录
        excel_name = Path(excel_path).stem
        excel_output_folder = Path(output_folder) / excel_name
        excel_output_folder.mkdir(parents=True, exist_ok=True)

        # 读取Excel数据
        df = pd.read_excel(excel_path, header=None, engine='openpyxl')

        # 检查列数是否足够
        if len(df.columns) < 6:
            raise ValueError("Excel文件必须包含至少6列数据（A-F列）")

        # 遍历每个模板文件
        for template_path in template_paths:
            # 只处理.inp后缀的文件
            if template_path.lower().endswith('.inp'):
                # 创建基于INP模板文件名的子文件夹
                template_name = Path(template_path).stem
                template_output_folder = excel_output_folder / template_name
                template_output_folder.mkdir(parents=True, exist_ok=True)

                # 遍历每一行数据
                for index, row in df.iterrows():
                    # 提取数据并校验类型
                    try:
                        fx = int(row[0])
                        fy = int(row[1])
                        fz = int(row[2])
                        rx = float(row[3])
                        ry = float(row[4])
                        rz = float(row[5])
                    except (ValueError, TypeError) as e:
                        raise ValueError(f"第{index + 1}行数据类型错误: {str(e)}")

                    # 读取原始模板
                    with open(template_path, 'r', encoding='utf-8') as f:
                        template_content = f.read()

                    # 构建替换内容
                    boundary_block = "\n".join([
                        "Set-1, 4, 4, {:.3f}".format(rx),
                        "Set-1, 5, 5, {:.3f}".format(ry),
                        "Set-1, 6, 6, {:.3f}".format(rz)
                    ])

                    cload_block = "\n".join([
                        "Set-1, 1, {}".format(fx),
                        "Set-1, 2, {}".format(fy),
                        "Set-1, 3, {}".format(fz)
                    ])

                    # 正则表达式精准匹配
                    boundary_pattern = re.compile(
                        r'\*Boundary, op=NEW, amplitude=Amp-1\n'
                        r'Set-\d+, \d+, \d+,[\d\.\-]+\n'
                        r'Set-\d+, \d+, \d+,[\d\.\-]+\n'
                        r'Set-\d+, \d+, \d+,[\d\.\-]+',
                        re.MULTILINE
                    )

                    cload_pattern = re.compile(
                        r'\*Cload, amplitude=Amp-1\n'
                        r'Set-\d+, \d+,[\d\-]+\n'
                        r'Set-\d+, \d+,[\d\-]+\n'
                        r'Set-\d+, \d+,[\d\-]+',
                        re.MULTILINE
                    )

                    # 执行替换
                    new_content = boundary_pattern.sub(
                        r'*Boundary, op=NEW, amplitude=Amp-1\n' + boundary_block,
                        template_content
                    )
                    new_content = cload_pattern.sub(
                        r'*Cload, amplitude=Amp-1\n' + cload_block,
                        new_content
                    )

                    # 写入文件
                    output_path = template_output_folder / f"input_{index + 1}.inp"
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)

                    print(f"成功生成文件：{output_path}")

    except Exception as e:
        raise RuntimeError(f"生成失败: {str(e)}")


class INPGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("INP文件批量生成器")
        self.create_widgets()

    def create_widgets(self):
        # 模板文件夹选择
        ttk.Label(self.master, text="模板INP文件夹:").grid(row=0, column=0, padx=5, pady=5)
        self.inp_folder_entry = ttk.Entry(width=40)
        self.inp_folder_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="浏览...", command=self.select_inp_folder).grid(row=0, column=2, padx=5, pady=5)

        # Excel文件选择（支持多选）
        ttk.Label(self.master, text="数据Excel文件:").grid(row=1, column=0, padx=5, pady=5)
        self.excel_entry = ttk.Entry(width=40)
        self.excel_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="浏览...", command=self.select_excel).grid(row=1, column=2, padx=5, pady=5)

        # 输出路径选择
        ttk.Label(self.master, text="输出路径:").grid(row=2, column=0, padx=5, pady=5)
        self.output_entry = ttk.Entry(width=40)
        self.output_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(self.master, text="浏览...", command=self.select_output).grid(row=2, column=2, padx=5, pady=5)

        # 生成按钮
        ttk.Button(self.master, text="开始生成", command=self.start_generation).grid(row=3, column=1, pady=10)

        # 日志文本框
        self.log = tk.Text(self.master, height=10, width=50)
        self.log.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

    def select_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)

    def select_inp_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.inp_folder_entry.delete(0, tk.END)
            self.inp_folder_entry.insert(0, folder)

    def select_excel(self):
        filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if filepaths:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, ";".join(filepaths))

    def start_generation(self):
        inp_folder = self.inp_folder_entry.get()
        excel_paths = self.excel_entry.get()
        output_folder = self.output_entry.get()

        if not all([inp_folder, excel_paths, output_folder]):
            messagebox.showerror("错误", "请先选择模板文件夹、Excel文件和输出路径")
            return

        # 获取模板文件夹下所有INP文件
        template_files = []
        if Path(inp_folder).is_dir():
            for file in Path(inp_folder).iterdir():
                if file.is_file() and file.suffix.lower() == '.inp':
                    template_files.append(str(file))

        if not template_files:
            messagebox.showerror("错误", f"模板文件夹中未找到INP文件: {inp_folder}")
            return

        # 分割多个Excel文件路径
        excel_files = excel_paths.split(";")

        # 检查Excel文件是否存在
        invalid_files = [f for f in excel_files if not Path(f).exists()]
        if invalid_files:
            messagebox.showerror("错误", f"以下Excel文件不存在:\n{', '.join(invalid_files)}")
            return

        try:
            # 显示选择的模板文件
            self.log.insert(tk.END, f"找到 {len(template_files)} 个模板文件:\n")
            for template in template_files:
                self.log.insert(tk.END, f" - {Path(template).name}\n")
            self.log.insert(tk.END, "\n")
            self.log.see(tk.END)

            for excel_file in excel_files:
                self.log.insert(tk.END, f"开始处理Excel文件: {excel_file}\n")
                self.log.see(tk.END)

                # 对每个Excel文件，处理所有模板
                generate_inps(excel_file, template_files, output_folder)

                self.log.insert(tk.END, f"✓ Excel文件处理完成: {excel_file}\n")
                self.log.see(tk.END)

            self.log.insert(tk.END, f"\n所有文件生成成功！保存路径：{output_folder}\n")
            self.log.see(tk.END)

        except Exception as e:
            self.log.insert(tk.END, f"错误: {str(e)}\n")
            self.log.see(tk.END)
            messagebox.showerror("错误", f"生成失败: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = INPGeneratorApp(root)
    root.mainloop()