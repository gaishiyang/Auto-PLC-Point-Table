# -*- coding: utf-8 -*-
"""EPLAN模块点位表生成器 - GUI界面"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from generate_eplan import generate_eplan_sheet


class EPLANGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EPLAN模块点位表生成器")
        self.root.geometry("500x450")
        
        main_frame = ttk.Frame(root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="EPLAN模块点位表生成器", font=("Arial", 14, "bold")).pack(pady=5)
        
        # I模块设置
        i_frame = ttk.LabelFrame(main_frame, text="I模块设置", padding="10")
        i_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(i_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.i_module = tk.IntVar(value=4)
        ttk.Spinbox(i_frame, from_=1, to=32, textvariable=self.i_module, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(i_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.i_start = tk.IntVar(value=2)
        ttk.Spinbox(i_frame, from_=0, to=1000, textvariable=self.i_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(i_frame, text="(共 0 点)").grid(row=0, column=4, sticky=tk.W)
        i_frame.grid_columnconfigure(4, weight=1)
        
        # Q模块设置
        q_frame = ttk.LabelFrame(main_frame, text="Q模块设置", padding="10")
        q_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(q_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.q_module = tk.IntVar(value=4)
        ttk.Spinbox(q_frame, from_=1, to=32, textvariable=self.q_module, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(q_frame, text="起始字节:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.q_start = tk.IntVar(value=2)
        ttk.Spinbox(q_frame, from_=0, to=1000, textvariable=self.q_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(q_frame, text="(共 0 点)").grid(row=0, column=4, sticky=tk.W)
        q_frame.grid_columnconfigure(4, weight=1)
        
        # IW模块设置
        iw_frame = ttk.LabelFrame(main_frame, text="IW模块设置", padding="10")
        iw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(iw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.iw_module = tk.IntVar(value=2)
        ttk.Spinbox(iw_frame, from_=1, to=32, textvariable=self.iw_module, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(iw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.iw_start = tk.IntVar(value=10)
        ttk.Spinbox(iw_frame, from_=0, to=1000, textvariable=self.iw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(iw_frame, text="(共 0 点)").grid(row=0, column=4, sticky=tk.W)
        iw_frame.grid_columnconfigure(4, weight=1)
        
        # QW模块设置
        qw_frame = ttk.LabelFrame(main_frame, text="QW模块设置", padding="10")
        qw_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(qw_frame, text="模块数量:").grid(row=0, column=0, sticky=tk.W)
        self.qw_module = tk.IntVar(value=2)
        ttk.Spinbox(qw_frame, from_=1, to=32, textvariable=self.qw_module, width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(qw_frame, text="起始偏移:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.qw_start = tk.IntVar(value=16)
        ttk.Spinbox(qw_frame, from_=0, to=1000, textvariable=self.qw_start, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(qw_frame, text="(共 0 点)").grid(row=0, column=4, sticky=tk.W)
        qw_frame.grid_columnconfigure(4, weight=1)
        
        # 输出文件
        file_frame = ttk.LabelFrame(main_frame, text="输出设置", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="输出文件:").grid(row=0, column=0, sticky=tk.W)
        self.output_path = tk.StringVar(value="EPLAN点位表.xlsx")
        ttk.Entry(file_frame, textvariable=self.output_path, width=35).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.browse_output).grid(row=0, column=2)
        
        # 按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="生成", command=self.generate, width=15).pack()
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(5, 0))
    
    def browse_output(self):
        filename = filedialog.asksaveasfilename(title="保存位置", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if filename:
            self.output_path.set(filename)
    
    def generate(self):
        try:
            output_file = self.output_path.get()
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            self.status_var.set("正在生成...")
            
            generate_eplan_sheet(
                module_i=self.i_module.get(),
                start_i=self.i_start.get(),
                module_q=self.q_module.get(),
                start_q=self.q_start.get(),
                module_iw=self.iw_module.get(),
                start_iw=self.iw_start.get(),
                module_qw=self.qw_module.get(),
                start_qw=self.qw_start.get(),
                output_file=output_file
            )
            
            self.status_var.set("生成完成!")
            messagebox.showinfo("成功", f"EPLAN点位表已生成!\n\n文件: {output_file}")
        except Exception as e:
            self.status_var.set(f"错误: {str(e)}")
            messagebox.showerror("错误", f"生成失败:\n{str(e)}")


if __name__ == '__main__':
    root = tk.Tk()
    app = EPLANGeneratorApp(root)
    root.mainloop()
