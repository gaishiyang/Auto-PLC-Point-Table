# -*- coding: utf-8 -*-
"""PLC点位表生成器 - 整合版"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_plc import parse_template, expand_points, octal_bit, octal_word, generate_var_name, generate_alias, export_alias_map, load_alias_map
from generate_module import generate_module_excel, generate_i_modules, generate_q_modules, generate_iw_modules, generate_qw_modules
from generate_terminal import generate_terminal_excel, parse_plc_points, group_devices_by_type, generate_terminal_rows_with_spacing, load_priority_table, load_priority_json, DEFAULT_PRIORITY


class PLCGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PLC点位表生成器")
        self.root.geometry("800x840")
        
        # 创建Notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建标签页
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="点位表生成")
        self.notebook.add(self.tab2, text="模块排序生成")
        self.notebook.add(self.tab3, text="端子表生成")
        
        # 初始化三个界面
        self.init_point_tab()
        self.init_module_tab()
        self.init_terminal_tab()
    
    # ========== 点位表生成标签页 ==========
    def init_point_tab(self):
        main_frame = ttk.Frame(self.tab1, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="PLC点位表生成", font=("Arial", 12, "bold")).pack(pady=5)
        
        file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding="8")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="模板文件:").grid(row=0, column=0, sticky=tk.W)
        self.template_path = tk.StringVar(value="点表模板.xlsx")
        ttk.Entry(file_frame, textvariable=self.template_path, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_template).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        self.output_path = tk.StringVar(value="PLC点位表.xlsx")
        ttk.Entry(file_frame, textvariable=self.output_path, width=40).grid(row=1, column=1, padx=5, pady=(5,0))
        ttk.Button(file_frame, text="浏览", command=self.browse_output).grid(row=1, column=2, pady=(5,0))
        
        # 别名映射表设置
        alias_frame = ttk.LabelFrame(main_frame, text="别名映射表(可选)", padding="8")
        alias_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(alias_frame, text="别名映射文件:").grid(row=0, column=0, sticky=tk.W)
        self.alias_file = tk.StringVar()
        ttk.Entry(alias_frame, textvariable=self.alias_file, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(alias_frame, text="导入", command=self.import_alias).grid(row=0, column=2, padx=2)
        ttk.Button(alias_frame, text="导出默认", command=self.export_alias).grid(row=0, column=3, padx=2)
        
        ttk.Label(alias_frame, text="(自定义设备类型对应的别名模板，如D_电机.设备.点位)", foreground="gray").grid(
            row=1, column=0, columnspan=4, sticky=tk.W, pady=(3,0))
        
        # 优先级表设置
        priority_frame = ttk.LabelFrame(main_frame, text="优先级表(可选，用于排序规则)", padding="8")
        priority_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(priority_frame, text="优先级文件:").grid(row=0, column=0, sticky=tk.W)
        self.point_priority_file = tk.StringVar()
        ttk.Entry(priority_frame, textvariable=self.point_priority_file, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(priority_frame, text="导入", command=self.import_priority_point).grid(row=0, column=2, padx=2)
        ttk.Button(priority_frame, text="导出默认", command=self.export_priority_point).grid(row=0, column=3, padx=2)
        
        ttk.Label(priority_frame, text="(自定义关键词排序优先级，数字越小越靠前，如远程=0,本地=1)", foreground="gray").grid(
            row=1, column=0, columnspan=4, sticky=tk.W, pady=(3,0))
        
        # 按钮（放在优先级表和起始地址之间）
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="刷新预览", command=self.refresh_preview).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成点位表", command=self.generate_plc, width=15).pack(side=tk.LEFT, padx=5)
        
        addr_frame = ttk.LabelFrame(main_frame, text="起始地址设置", padding="8")
        addr_frame.pack(fill=tk.X, pady=5)
        
        self.i_start = tk.IntVar(value=2)
        self.q_start = tk.IntVar(value=0)
        self.iw_start = tk.IntVar(value=10)
        self.qw_start = tk.IntVar(value=0)
        
        rows = [
            ("I地址起始字节:", self.i_start, "I字节.位"),
            ("Q地址起始字节:", self.q_start, "Q字节.位"),
            ("IW地址起始偏移:", self.iw_start, "IW偏移"),
            ("QW地址起始偏移:", self.qw_start, "QW偏移"),
        ]
        
        for i, (label, var, note) in enumerate(rows):
            ttk.Label(addr_frame, text=label).grid(row=i, column=0, sticky=tk.W, pady=3)
            ttk.Spinbox(addr_frame, from_=0, to=1000, textvariable=var, width=8).grid(row=i, column=1, sticky=tk.W)
            ttk.Label(addr_frame, text=f"({note})", foreground="gray").grid(row=i, column=2, sticky=tk.W, padx=5)
        
        preview_frame = ttk.LabelFrame(main_frame, text="设备类型预览", padding="8")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("设备类型", "设备数", "设备名称", "I", "Q", "IW", "QW")
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=8)
        
        widths = [80, 60, 150, 80, 80, 60, 60]
        for col, width in zip(columns, widths):
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=width, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=scrollbar.set)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.point_status = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.point_status, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(5, 0))
        
        self.refresh_preview()
    
    def browse_template(self):
        filename = filedialog.askopenfilename(title="选择模板", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.template_path.set(filename)
            self.refresh_preview()
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(title="保存位置", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.output_path.set(filename)
    
    def import_alias(self):
        filename = filedialog.askopenfilename(title="导入别名映射表", filetypes=[("Excel", "*.xlsx")])
        if filename:
            try:
                load_alias_map(filename)
                self.alias_file.set(filename)
                messagebox.showinfo("成功", f"已导入别名映射表:\n{filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败:\n{str(e)}")
    
    def export_alias(self):
        filename = filedialog.asksaveasfilename(title="导出别名映射表", defaultextension=".xlsx",
                                                initialfile="别名映射表.xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            try:
                export_alias_map(filename)
                self.alias_file.set(filename)
                messagebox.showinfo("成功", f"已导出别名映射表:\n{filename}\n\n可编辑后导入使用")
            except Exception as e:
                messagebox.showerror("错误", str(e))
    
    def import_priority_point(self):
        filename = filedialog.askopenfilename(title="导入优先级表", filetypes=[("Excel", "*.xlsx")])
        if filename:
            try:
                from generate_terminal import load_priority_table, save_priority_json
                priority_dict = load_priority_table(filename)
                if priority_dict:
                    save_priority_json(priority_dict)  # 保存到JSON供端子表使用
                self.point_priority_file.set(filename)
                messagebox.showinfo("成功", f"已导入优先级表:\n{filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导入失败:\n{str(e)}")
    
    def export_priority_point(self):
        from generate_terminal import export_priority_table
        filename = filedialog.asksaveasfilename(title="导出优先级表", defaultextension=".xlsx",
                                                initialfile="优先级表.xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            try:
                export_priority_table(filename)
                self.point_priority_file.set(filename)
                messagebox.showinfo("成功", f"已导出优先级表:\n{filename}\n\n可编辑后导入使用")
            except Exception as e:
                messagebox.showerror("错误", str(e))
    
    def refresh_preview(self):
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        path = self.template_path.get()
        if not os.path.exists(path):
            self.point_status.set(f"文件不存在: {path}")
            return
        
        try:
            device_definitions = parse_template(path)
            total_i = total_q = total_iw = total_qw = 0
            
            for dtype, defs in device_definitions.items():
                count = len(defs['devices'])
                i_cnt = count * len(defs['i_templates'])
                q_cnt = count * len(defs['q_templates'])
                iw_cnt = count * len(defs['iw_templates'])
                qw_cnt = count * len(defs['qw_templates'])
                
                total_i += i_cnt
                total_q += q_cnt
                total_iw += iw_cnt
                total_qw += qw_cnt
                
                devices_str = ",".join(defs['devices'][:4])
                if len(defs['devices']) > 4:
                    devices_str += "..."
                
                self.preview_tree.insert("", tk.END, values=(
                    dtype, count, devices_str,
                    ",".join(defs['i_templates']) if defs['i_templates'] else "-",
                    ",".join(defs['q_templates']) if defs['q_templates'] else "-",
                    ",".join(defs['iw_templates']) if defs['iw_templates'] else "-",
                    ",".join(defs['qw_templates']) if defs['qw_templates'] else "-"
                ))
            
            self.preview_tree.insert("", tk.END, values=(f"总计", f"I:{total_i} Q:{total_q} IW:{total_iw} QW:{total_qw}", "", "", "", "", ""))
            self.point_status.set(f"就绪 - 共 {len(device_definitions)} 种设备类型")
        except Exception as e:
            self.point_status.set(f"错误: {str(e)}")
    
    def generate_plc(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment
        
        template_file = self.template_path.get()
        output_file = self.output_path.get()
        
        if not os.path.exists(template_file):
            messagebox.showerror("错误", "模板文件不存在!")
            return
        
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
        
        # 获取别名映射表
        alias_file = self.alias_file.get() if self.alias_file.get() else None
        if alias_file and not os.path.exists(alias_file):
            alias_file = None
        alias_map = load_alias_map(alias_file)
        
        try:
            self.point_status.set("正在生成...")
            
            device_definitions = parse_template(template_file)
            i_points, q_points, iw_points, qw_points = expand_points(device_definitions)
            
            wb = Workbook()
            ws = wb.active
            ws.title = "PLC点位"
            
            headers = ['I地址', 'I变量名', 'I别名', 'Q地址', 'Q变量名', 'Q别名', 'IW地址', 'IW变量名', 'IW别名', 'QW地址', 'QW变量名', 'QW别名']
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            
            i_start = self.i_start.get()
            q_start = self.q_start.get()
            iw_start = self.iw_start.get()
            qw_start = self.qw_start.get()
            
            for idx, point in enumerate(i_points):
                ws.cell(row=idx+2, column=1, value=octal_bit('I', i_start, idx))
                ws.cell(row=idx+2, column=2, value=generate_var_name(point))
                ws.cell(row=idx+2, column=3, value=generate_alias(point, alias_map))
            
            for idx, point in enumerate(q_points):
                ws.cell(row=idx+2, column=4, value=octal_bit('Q', q_start, idx))
                ws.cell(row=idx+2, column=5, value=generate_var_name(point))
                ws.cell(row=idx+2, column=6, value=generate_alias(point, alias_map))
            
            for idx, point in enumerate(iw_points):
                ws.cell(row=idx+2, column=7, value=octal_word('IW', iw_start, idx, 2))
                ws.cell(row=idx+2, column=8, value=generate_var_name(point))
                ws.cell(row=idx+2, column=9, value=generate_alias(point, alias_map))
            
            for idx, point in enumerate(qw_points):
                ws.cell(row=idx+2, column=10, value=octal_word('QW', qw_start, idx, 2))
                ws.cell(row=idx+2, column=11, value=generate_var_name(point))
                ws.cell(row=idx+2, column=12, value=generate_alias(point, alias_map))
            
            for col_idx in range(1, 13):
                max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) 
                              for r in range(1, max(len(i_points), len(q_points), len(iw_points), len(qw_points)) + 2))
                ws.column_dimensions[chr(64 + col_idx) if col_idx <= 26 else 'A' + chr(64 + col_idx - 26)].width = max_len + 3
            
            wb.save(output_file)
            
            alias_info = f"\n别名映射: 自定义" if alias_file else "\n别名映射: 默认"
            self.point_status.set("生成完成!")
            messagebox.showinfo("成功", f"PLC点位表已生成!{alias_info}\n\n文件: {output_file}\n\nI: {len(i_points)} 个\nQ: {len(q_points)} 个\nIW: {len(iw_points)} 个\nQW: {len(qw_points)} 个")
        except Exception as e:
            self.point_status.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"生成失败:\n{str(e)}")
    
    # ========== 模块排序生成标签页 ==========
    def init_module_tab(self):
        main_frame = ttk.Frame(self.tab2, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="PLC模块排序生成", font=("Arial", 12, "bold")).pack(pady=5)
        
        i_frame = ttk.LabelFrame(main_frame, text="I模块设置", padding="8")
        i_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(i_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.mod_i_count = tk.IntVar(value=4)
        ttk.Spinbox(i_frame, from_=1, to=32, textvariable=self.mod_i_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(i_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.mod_i_start = tk.IntVar(value=2)
        ttk.Spinbox(i_frame, from_=0, to=1000, textvariable=self.mod_i_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(i_frame, text=f"(共 {self.mod_i_count.get()*16} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        
        q_frame = ttk.LabelFrame(main_frame, text="Q模块设置", padding="8")
        q_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(q_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.mod_q_count = tk.IntVar(value=4)
        ttk.Spinbox(q_frame, from_=1, to=32, textvariable=self.mod_q_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(q_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.mod_q_start = tk.IntVar(value=2)
        ttk.Spinbox(q_frame, from_=0, to=1000, textvariable=self.mod_q_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(q_frame, text=f"(共 {self.mod_q_count.get()*16} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        
        iw_frame = ttk.LabelFrame(main_frame, text="IW模块设置", padding="8")
        iw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(iw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.mod_iw_count = tk.IntVar(value=2)
        ttk.Spinbox(iw_frame, from_=1, to=32, textvariable=self.mod_iw_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(iw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.mod_iw_start = tk.IntVar(value=10)
        ttk.Spinbox(iw_frame, from_=0, to=1000, textvariable=self.mod_iw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(iw_frame, text=f"(共 {self.mod_iw_count.get()*8} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        
        qw_frame = ttk.LabelFrame(main_frame, text="QW模块设置", padding="8")
        qw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(qw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.mod_qw_count = tk.IntVar(value=2)
        ttk.Spinbox(qw_frame, from_=1, to=32, textvariable=self.mod_qw_count, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(qw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20,0))
        self.mod_qw_start = tk.IntVar(value=16)
        ttk.Spinbox(qw_frame, from_=0, to=1000, textvariable=self.mod_qw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(qw_frame, text=f"(共 {self.mod_qw_count.get()*4} 点)", foreground="gray").grid(row=0, column=4, sticky=tk.W)
        
        # PLC类型选择
        type_frame = ttk.LabelFrame(main_frame, text="PLC类型", padding="8")
        type_frame.pack(fill=tk.X, pady=5)
        
        self.mod_plc_type = tk.StringVar(value="200SMART")
        ttk.Radiobutton(type_frame, text="S7-1500", variable=self.mod_plc_type, value="1500", command=self.module_preview).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_frame, text="200SMART", variable=self.mod_plc_type, value="200SMART", command=self.module_preview).pack(side=tk.LEFT, padx=10)
        
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(out_frame, text="输出文件:").pack(side=tk.LEFT)
        self.module_output = tk.StringVar(value="PLC模块排序表.xlsx")
        ttk.Entry(out_frame, textvariable=self.module_output, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(out_frame, text="浏览", command=self.browse_module_output).pack(side=tk.LEFT)
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="预览", command=self.module_preview).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成排序表", command=self.generate_module, width=15).pack(side=tk.LEFT, padx=5)
        
        preview_frame = ttk.LabelFrame(main_frame, text="预览", padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.module_preview_text = tk.Text(preview_frame, height=8, font=("Consolas", 9))
        self.module_preview_text.pack(fill=tk.BOTH, expand=True)
        
        self.module_preview()
    
    def browse_module_output(self):
        filename = filedialog.asksaveasfilename(title="保存位置", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.module_output.set(filename)
    
    def module_preview(self):
        plc_type = self.mod_plc_type.get()
        i_pts = generate_i_modules(self.mod_i_count.get(), self.mod_i_start.get())
        q_pts = generate_q_modules(self.mod_q_count.get(), self.mod_q_start.get())
        iw_pts = generate_iw_modules(self.mod_iw_count.get(), self.mod_iw_start.get(), plc_type)
        qw_pts = generate_qw_modules(self.mod_qw_count.get(), self.mod_qw_start.get(), plc_type)
        
        preview = f"""【PLC类型】{plc_type}

【I模块】{len(i_pts)}点 ({self.mod_i_count.get()}模块，起始I{self.mod_i_start.get()}.0)
  {', '.join(i_pts[:8])}...

【Q模块】{len(q_pts)}点 ({self.mod_q_count.get()}模块，起始Q{self.mod_q_start.get()}.0)
  {', '.join(q_pts[:8])}...

【IW模块】{len(iw_pts)}点 ({self.mod_iw_count.get()}模块，起始IW{self.mod_iw_start.get()})
  {', '.join(iw_pts[:16])}{'...' if len(iw_pts) > 16 else ''}

【QW模块】{len(qw_pts)}点 ({self.mod_qw_count.get()}模块，起始QW{self.mod_qw_start.get()})
  {', '.join(qw_pts)}"""
        
        self.module_preview_text.delete(1.0, tk.END)
        self.module_preview_text.insert(1.0, preview)
    
    def generate_module(self):
        try:
            output = self.module_output.get()
            if not output.endswith('.xlsx'):
                output += '.xlsx'
            
            generate_module_excel(
                self.mod_i_count.get(), self.mod_i_start.get(),
                self.mod_q_count.get(), self.mod_q_start.get(),
                self.mod_iw_count.get(), self.mod_iw_start.get(),
                self.mod_qw_count.get(), self.mod_qw_start.get(),
                output,
                plc_type=self.mod_plc_type.get()
            )
            messagebox.showinfo("成功", f"已生成: {output}")
        except Exception as e:
            messagebox.showerror("错误", str(e))
    
    # ========== 端子表生成标签页 ==========
    def init_terminal_tab(self):
        main_frame = ttk.Frame(self.tab3, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="端子表生成", font=("Arial", 12, "bold")).pack(pady=5)
        
        # 文件设置
        file_frame = ttk.LabelFrame(main_frame, text="文件设置", padding="8")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="PLC点位表:").grid(row=0, column=0, sticky=tk.W)
        self.terminal_input = tk.StringVar(value="PLC点位表.xlsx")
        ttk.Entry(file_frame, textvariable=self.terminal_input, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_terminal_input).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        self.terminal_output = tk.StringVar(value="端子表_生成.xlsx")
        ttk.Entry(file_frame, textvariable=self.terminal_output, width=40).grid(row=1, column=1, padx=5, pady=(5,0))
        ttk.Button(file_frame, text="浏览", command=self.browse_terminal_output).grid(row=1, column=2, pady=(5,0))
        
        # PLC类型和24V设置
        setting_frame = ttk.LabelFrame(main_frame, text="PLC类型与电源设置", padding="8")
        setting_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(setting_frame, text="PLC类型:").grid(row=0, column=0, sticky=tk.W)
        self.plc_type = tk.StringVar(value="200SMART")
        ttk.Radiobutton(setting_frame, text="S7-1500", variable=self.plc_type, value="1500").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(setting_frame, text="200SMART", variable=self.plc_type, value="200SMART").grid(row=0, column=2, padx=5)
        
        ttk.Label(setting_frame, text="24V+标识:").grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        self.power_pos = tk.StringVar(value="24V+")
        ttk.Entry(setting_frame, textvariable=self.power_pos, width=15).grid(row=1, column=1, padx=5, pady=(5,0), sticky=tk.W)
        
        ttk.Label(setting_frame, text="24V-标识:").grid(row=1, column=2, sticky=tk.W, pady=(5,0))
        self.power_neg = tk.StringVar(value="24V-")
        ttk.Entry(setting_frame, textvariable=self.power_neg, width=15).grid(row=1, column=3, padx=5, pady=(5,0), sticky=tk.W)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="刷新预览", command=self.terminal_preview).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="生成端子表", command=self.generate_terminal, width=15).pack(side=tk.LEFT, padx=5)
        
        # 设备分组预览
        preview_frame = ttk.LabelFrame(main_frame, text="设备分组预览", padding="8")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        columns = ("设备名", "I点数", "Q点数", "IW点数", "QW点数", "Q地址")
        self.terminal_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=10)
        
        widths = [80, 60, 60, 60, 60, 120]
        for col, width in zip(columns, widths):
            self.terminal_tree.heading(col, text=col)
            self.terminal_tree.column(col, width=width, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.terminal_tree.yview)
        self.terminal_tree.configure(yscrollcommand=scrollbar.set)
        
        self.terminal_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.terminal_status = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.terminal_status, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(5, 0))
        
        self.terminal_preview()
    
    def browse_terminal_input(self):
        filename = filedialog.askopenfilename(title="选择PLC点位表", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.terminal_input.set(filename)
            self.terminal_preview()
    
    def browse_terminal_output(self):
        filename = filedialog.asksaveasfilename(title="保存位置", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.terminal_output.set(filename)
    
    def terminal_preview(self):
        for item in self.terminal_tree.get_children():
            self.terminal_tree.delete(item)
        
        path = self.terminal_input.get()
        if not os.path.exists(path):
            self.terminal_status.set(f"文件不存在: {path}")
            return
        
        try:
            i_pts, q_pts, iw_pts, qw_pts = parse_plc_points(path)
            devices = group_devices_by_type(i_pts, q_pts, iw_pts, qw_pts)
            
            for device in sorted(devices.keys()):
                pts = devices[device]
                q_addrs = list(pts['Q'].values()) if pts['Q'] else ['-']
                self.terminal_tree.insert("", tk.END, values=(
                    device,
                    len(pts['I']),
                    len(pts['Q']),
                    len(pts['IW']),
                    len(pts['QW']),
                    ", ".join(q_addrs[:3])
                ))
            
            self.terminal_status.set(f"就绪 - 共 {len(devices)} 个设备")
        except Exception as e:
            self.terminal_status.set(f"错误: {str(e)}")
    
    def generate_terminal(self):
        input_file = self.terminal_input.get()
        output_file = self.terminal_output.get()
        
        if not os.path.exists(input_file):
            messagebox.showerror("错误", "PLC点位表不存在!")
            return
        
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
        
        # 获取PLC类型和24V设置
        plc_type = self.plc_type.get()
        power_pos = self.power_pos.get() or "24V+"
        power_neg = self.power_neg.get() or "24V-"
        
        try:
            self.terminal_status.set("正在生成...")
            generate_terminal_excel(input_file, output_file, plc_type=plc_type, power_pos=power_pos, power_neg=power_neg)
            self.terminal_status.set("生成完成!")
            has_custom = load_priority_json() is not None
            priority_info = "\n优先级: 自定义" if has_custom else "\n优先级: 默认"
            messagebox.showinfo("成功", f"端子表已生成!\nPLC类型: {plc_type}{priority_info}\n\n文件: {output_file}")
        except Exception as e:
            self.terminal_status.set(f"错误: {str(e)}")
            messagebox.showerror("错误", str(e))


if __name__ == '__main__':
    root = tk.Tk()
    app = PLCGeneratorApp(root)
    root.mainloop()
