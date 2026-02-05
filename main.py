import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import time

# ================= 核心配置区域 =================

# 扫描区域
SCAN_AREAS = [
    (6, 11, 'B'),
    (13, 17, 'C'),
    (19, 29, 'C'),
    (31, 33, 'C'),
    (35, 42, 'C'),
    (44, 55, 'C'),
    (57, 62, 'C'),
    (64, 72, 'C'),
    (74, 79, 'C'),
    (81, 95, 'C'),
    (97, 108, 'C'),
    (110, 134, 'B'),
    (136, 152, 'C'),
    (153, 164, 'B')
]

# 注意：这里删除了 DATA_COL_INDICES，改为动态计算

NAME_MAPPING = {
    '抗链球菌溶血素O测定': '抗链球菌溶血素 O 测定',
    'EB病毒三项': 'EB病毒',
    '尿白蛋白肌酐比（ACR）': '尿白蛋白肌酐比',
    '尿常规': '尿常规/沉渣',
    '头部MRA和头颅平扫': '头部MRA',
    '腰椎膝盖磁共振': '腰椎膝盖核磁',
    'X线': '胸部X线',
    '电子胃镜和电子肠镜（无痛）': '电子胃肠镜',
}

DATE_COL_NAME = "检查日期"
REMARK_COL_NAME = "备注_未匹配及模糊项"


# =========================================================

class HealthDataApp:
    def __init__(self, root):
        self.root = root
        self.root.title("体检报告自动转换工具")
        self.root.geometry("600x500")

        self.source_path = tk.StringVar()
        self.template_path = tk.StringVar()

        # === 界面布局 ===
        tk.Label(root, text="第一步：选择需要处理的文件").pack(pady=(10, 0))
        frame_src = tk.Frame(root)
        frame_src.pack(pady=5, padx=20, fill="x")
        tk.Entry(frame_src, textvariable=self.source_path, width=50).pack(side="left", padx=5)
        tk.Button(frame_src, text="浏览...", command=self.select_source).pack(side="left")

        tk.Label(root, text="第二步：选择模板文件 (标准空表)").pack(pady=(10, 0))
        frame_tmpl = tk.Frame(root)
        frame_tmpl.pack(pady=5, padx=20, fill="x")
        tk.Entry(frame_tmpl, textvariable=self.template_path, width=50).pack(side="left", padx=5)
        tk.Button(frame_tmpl, text="浏览...", command=self.select_template).pack(side="left")

        self.btn_run = tk.Button(root, text="开始转换", command=self.start_processing_thread, bg="#4CAF50", fg="white",
                                 font=("Arial", 12, "bold"))
        self.btn_run.pack(pady=20, ipadx=20, ipady=5)

        tk.Label(root, text="运行日志：").pack(anchor="w", padx=20)
        self.log_area = scrolledtext.ScrolledText(root, height=12, state='disabled', bg="#f0f0f0")
        self.log_area.pack(padx=20, pady=5, fill="both", expand=True)

    def log(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update()

    def select_source(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename: self.source_path.set(filename)

    def select_template(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filename: self.template_path.set(filename)

    def start_processing_thread(self):
        if not self.source_path.get() or not self.template_path.get():
            messagebox.showwarning("提示", "请先选择源文件和模板文件！")
            return
        self.btn_run.config(state='disabled', text="正在处理...")
        thread = threading.Thread(target=self.run_process)
        thread.start()

    def run_process(self):
        try:
            input_file = self.source_path.get()
            template_file = self.template_path.get()
            dir_name, base_name = os.path.split(input_file)
            name_only, ext = os.path.splitext(base_name)
            output_file = os.path.join(dir_name, f"{name_only}_填报结果{ext}")

            self.log("-" * 30)
            self.log("开始任务...")

            # 读取模板
            template_df = pd.read_excel(template_file, header=0)
            template_columns = list(template_df.columns)

            # 读取所有 Sheet
            self.log("正在读取所有 Sheet 数据...")
            all_sheets_dict = pd.read_excel(input_file, header=None, sheet_name=None)
            self.log(f"共发现 {len(all_sheets_dict)} 个 Sheet。")

            all_rows_to_append = []

            for sheet_name, df in all_sheets_dict.items():
                self.log(f"正在处理 Sheet: {sheet_name}")

                anchor_rows = []
                for idx, row in df.iterrows():
                    if len(row) < 5: continue
                    row_str = " ".join(row.iloc[0:5].astype(str).tolist())
                    if "姓名" in row_str and "出生日期" in row_str:
                        anchor_rows.append(idx)

                if not anchor_rows: continue

                for start_row in anchor_rows:
                    date_row_idx = start_row + 2
                    active_dates = []
                    max_col_index = df.shape[1]

                    # === 【关键修改】动态生成列索引 ===
                    # 从索引4(E列)开始，每隔2列读一次，直到表格最右侧
                    # 这样无论是6次、8次还是10次，都能覆盖到
                    for col_idx in range(4, max_col_index, 2):

                        val = df.iloc[date_row_idx, col_idx]
                        # 检查是否为空 (有效性检查)
                        if pd.notna(val) and str(val).strip() != "":
                            active_dates.append((col_idx, str(val).strip()))

                    if not active_dates: continue

                    for col_idx, date_str in active_dates:
                        new_row_data = {col: "" for col in template_columns}
                        new_row_data[DATE_COL_NAME] = date_str

                        unmatched_buffer = []
                        used_targets = set()

                        for area_start, area_end, label_col_char in SCAN_AREAS:
                            label_col_idx = ord(label_col_char.upper()) - 65

                            for offset in range(area_start, area_end + 1):
                                current_row = start_row + offset
                                if current_row >= len(df): break
                                if label_col_idx >= df.shape[1]: continue

                                raw_label = df.iloc[current_row, label_col_idx]
                                if col_idx >= df.shape[1]: continue
                                value = df.iloc[current_row, col_idx]

                                if pd.notna(raw_label) and str(raw_label).strip() != "":
                                    src_name = str(raw_label).strip()
                                    val_str = str(value).strip() if pd.notna(value) else ""
                                    if val_str == "": continue

                                    target_name = NAME_MAPPING.get(src_name, src_name)

                                    if target_name in template_columns:
                                        if target_name not in used_targets:
                                            new_row_data[target_name] = val_str
                                            used_targets.add(target_name)
                                        else:
                                            unmatched_buffer.append(f"{src_name}(重复):{val_str}")
                                    else:
                                        unmatched_buffer.append(f"{src_name}:{val_str}")

                        if unmatched_buffer:
                            new_row_data[REMARK_COL_NAME] = "；".join(unmatched_buffer)

                        all_rows_to_append.append(new_row_data)

            if all_rows_to_append:
                self.log(f"正在写入 Excel，共 {len(all_rows_to_append)} 条数据...")
                new_data_df = pd.DataFrame(all_rows_to_append)

                final_cols = list(template_columns)
                if DATE_COL_NAME in final_cols: final_cols.remove(DATE_COL_NAME)
                final_cols.insert(0, DATE_COL_NAME)

                if REMARK_COL_NAME in final_cols: final_cols.remove(REMARK_COL_NAME)
                if REMARK_COL_NAME in new_data_df.columns: final_cols.append(REMARK_COL_NAME)

                final_df = new_data_df.reindex(columns=final_cols)
                final_df.to_excel(output_file, index=False)

                self.log("✅ 成功！")
                self.log(f"结果已保存至: {output_file}")
                messagebox.showinfo("成功", f"处理完成！\n文件已保存为：\n{output_file}")
            else:
                self.log("❌ 警告：未提取到任何数据。")
                messagebox.showwarning("警告", "未提取到任何数据，请检查文件格式。")

        except Exception as e:
            self.log(f"❌ 错误: {str(e)}")
            messagebox.showerror("运行出错", f"发生了错误:\n{str(e)}")

        finally:
            self.btn_run.config(state='normal', text="开始转换")


if __name__ == '__main__':
    root = tk.Tk()
    app = HealthDataApp(root)
    root.mainloop()